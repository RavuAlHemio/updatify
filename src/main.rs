mod semaphore;
mod variant;


use std::borrow::Cow;
use std::io::Write;
use std::os::windows::ffi::OsStrExt;
use std::path::PathBuf;
use std::process::ExitCode;
use std::ptr::null_mut;
use std::sync::Arc;

use clap::Parser;
use windows::core::{BSTR, ComObjectInterface, Error, PCWSTR, Ref, Interface, implement, w};
use windows::Win32::System::Com::{
    CLSCTX_ALL, CLSIDFromProgID, COINIT_MULTITHREADED, CoCreateInstance, CoInitializeEx,
};
use windows::Win32::System::Shutdown::InitiateSystemShutdownW;
use windows::Win32::System::UpdateAgent::{
    IDownloadCompletedCallback, IDownloadCompletedCallback_Impl, IDownloadCompletedCallbackArgs,
    IDownloadJob, IDownloadProgressChangedCallback, IDownloadProgressChangedCallback_Impl,
    IDownloadProgressChangedCallbackArgs, IInstallationCompletedCallback,
    IInstallationCompletedCallback_Impl, IInstallationCompletedCallbackArgs, IInstallationJob,
    IInstallationProgressChangedCallback, IInstallationProgressChangedCallback_Impl,
    IInstallationProgressChangedCallbackArgs, IUpdateCollection, IUpdateServiceManager,
    IUpdateServiceManager2, IUpdateSession, InstallationImpact, InstallationRebootBehavior,
    OperationResultCode, asfAllowOnlineRegistration, asfAllowPendingRegistration,
    asfRegisterServiceWithAU, iiMinor, iiNormal, iiRequiresExclusiveHandling,
    irbAlwaysRequiresReboot, irbCanRequestReboot, irbNeverReboots, orcAborted, orcFailed,
    orcInProgress, orcNotStarted, orcSucceeded, orcSucceededWithErrors, ssManagedServer, ssOthers,
    ssWindowsUpdate,
};

use crate::semaphore::Semaphore;
use crate::variant::null_variant;


const MS_UPDATE_ID: &str = "7971f918-a847-4430-9279-4a52d1efe18d";


#[derive(Parser)]
struct Opts {
    #[arg(short, long)]
    pub ms_update_opt_in: bool,

    #[arg(short, long)]
    pub local: Option<PathBuf>,

    #[arg(short, long)]
    pub skip_wsus: bool,

    #[arg(short, long)]
    pub interactive: bool,

    #[arg(long)]
    pub reboot_when_done: bool,

    #[arg(long)]
    pub criteria: Option<String>,
}

#[implement(IDownloadProgressChangedCallback, IDownloadCompletedCallback)]
struct DownloadState {
    continuation: Arc<Semaphore>,
}
#[allow(non_snake_case)]
impl IDownloadProgressChangedCallback_Impl for DownloadState_Impl {
    fn Invoke(
        &self,
        download_job_ref: Ref<IDownloadJob>,
        callback_args_ref: Ref<IDownloadProgressChangedCallbackArgs>,
    ) -> Result<(), Error> {
        let _ = download_job_ref;
        let callback_args = callback_args_ref.ok()?;
        let progress = unsafe { callback_args.Progress() }?;
        let current_percent = unsafe { progress.CurrentUpdatePercentComplete() }?;
        let total_percent = unsafe { progress.PercentComplete() }?;
        print!("\rdownload {:3}% / {:3}%", current_percent, total_percent);
        Ok(())
    }
}
#[allow(non_snake_case)]
impl IDownloadCompletedCallback_Impl for DownloadState_Impl {
    fn Invoke(
        &self,
        download_job_ref: Ref<IDownloadJob>,
        callback_args_ref: Ref<IDownloadCompletedCallbackArgs>,
    ) -> Result<(), Error> {
        let _ = (download_job_ref, callback_args_ref);
        self.continuation.increment();
        Ok(())
    }
}

#[implement(IInstallationProgressChangedCallback, IInstallationCompletedCallback)]
struct InstallState {
    continuation: Arc<Semaphore>,
}
#[allow(non_snake_case)]
impl IInstallationProgressChangedCallback_Impl for InstallState_Impl {
    fn Invoke(
        &self,
        install_job_ref: Ref<IInstallationJob>,
        callback_args_ref: Ref<IInstallationProgressChangedCallbackArgs>,
    ) -> Result<(), Error> {
        let _ = install_job_ref;
        let callback_args = callback_args_ref.ok()?;
        let progress = unsafe { callback_args.Progress() }?;
        let current_percent = unsafe { progress.CurrentUpdatePercentComplete() }?;
        let total_percent = unsafe { progress.PercentComplete() }?;
        print!("\rinstall  {:3}% / {:3}%", current_percent, total_percent);
        Ok(())
    }
}
#[allow(non_snake_case)]
impl IInstallationCompletedCallback_Impl for InstallState_Impl {
    fn Invoke(
        &self,
        install_job_ref: Ref<IInstallationJob>,
        callback_args_ref: Ref<IInstallationCompletedCallbackArgs>,
    ) -> Result<(), Error> {
        let _ = (install_job_ref, callback_args_ref);
        self.continuation.increment();
        Ok(())
    }
}


#[allow(non_upper_case_globals)]
fn result_code_string(result_code: OperationResultCode) -> Cow<'static, str> {
    match result_code {
        orcNotStarted => Cow::Borrowed("not started"),
        orcInProgress => Cow::Borrowed("in progress"),
        orcSucceeded => Cow::Borrowed("succeeded"),
        orcSucceededWithErrors => Cow::Borrowed("succeeded with errors"),
        orcFailed => Cow::Borrowed("failed"),
        orcAborted => Cow::Borrowed("aborted"),
        other => Cow::Owned(format!("unknown status code {}", other.0)),
    }
}

#[allow(non_upper_case_globals)]
fn minor_tag(impact: InstallationImpact) -> Cow<'static, str> {
    match impact {
        iiNormal => Cow::Borrowed(""),
        iiMinor => Cow::Borrowed(" [minor]"),
        iiRequiresExclusiveHandling => Cow::Borrowed(" [exclusive]"),
        other => Cow::Owned(format!(" [impact {}]", other.0)),
    }
}

#[allow(non_upper_case_globals)]
fn reboot_tag(reboot_behavior: InstallationRebootBehavior) -> Cow<'static, str> {
    match reboot_behavior {
        irbNeverReboots => Cow::Borrowed(" [bootless]"),
        irbAlwaysRequiresReboot => Cow::Borrowed(""),
        irbCanRequestReboot => Cow::Borrowed("[mayboot]"),
        other => Cow::Owned(format!("[reboot {}]", other.0)),
    }
}


fn main() -> ExitCode {
    let opts = Opts::parse();
    let mut exit_code = ExitCode::SUCCESS;

    // initialize COM
    unsafe {
        CoInitializeEx(
            None,
            COINIT_MULTITHREADED,
        )
    }
        .ok().expect("failed to initialize COM");

    let update_service_manager_guid = unsafe {
        CLSIDFromProgID(w!("Microsoft.Update.ServiceManager"))
    }
        .expect("failed to find class Microsoft.Update.ServiceManager");
    let update_service_manager: IUpdateServiceManager = unsafe {
        CoCreateInstance(
            &update_service_manager_guid,
            None,
            CLSCTX_ALL,
        )
    }
        .expect("failed to create update service manager");

    let mut offline_service_opt = None;
    if opts.ms_update_opt_in {
        let manager2: IUpdateServiceManager2 = update_service_manager.cast()
            .expect("failed to cast IUpdateServiceManager to IUpdateServiceManager2");
        unsafe {
            manager2.AddService2(
                &BSTR::from(MS_UPDATE_ID),
                asfAllowOnlineRegistration.0
                    | asfAllowPendingRegistration.0
                    | asfRegisterServiceWithAU.0,
                &BSTR::from(""),
            )
        }
            .expect("failed to register with Microsoft Update");
    }

    if let Some(loc) = opts.local {
        let loc_u16: Vec<u16> = loc.as_os_str().encode_wide().collect();
        let offline_service = unsafe {
            update_service_manager.AddScanPackageService(
                &BSTR::from("Offline Sync Service"),
                &BSTR::from_wide(&loc_u16),
                0,
            )
        }
            .expect("failed to add offline .cab for scanning");
        offline_service_opt = Some(offline_service);
    }

    // create an update session
    let update_session_guid = unsafe {
        CLSIDFromProgID(w!("Microsoft.Update.Session"))
    }
        .expect("failed to find class Microsoft.Update.Session");
    let update_session: IUpdateSession = unsafe {
        CoCreateInstance(
            &update_session_guid,
            None,
            CLSCTX_ALL,
        )
    }
        .expect("failed to create update session");

    // set up the searcher
    let update_searcher = unsafe {
        update_session.CreateUpdateSearcher()
    }
        .expect("failed to create update searcher");
    // only take a reference
    // (we can only afford dropping the offline service once installation is done)
    if let Some(offline_service) = offline_service_opt.as_ref() {
        unsafe {
            update_searcher.SetServerSelection(ssOthers)
        }
            .expect("failed to set service selection to Others for offline search");
        let service_id = unsafe {
            offline_service.ServiceID()
        }
            .expect("failed to obtain service ID for offline search");
        unsafe {
            update_searcher.SetServiceID(&service_id)
        }
            .expect("failed to set service ID for offline search");
    } else if opts.skip_wsus {
        unsafe {
            update_searcher.SetServerSelection(ssWindowsUpdate)
        }
            .expect("failed to set service selection to WindowsUpdate to skip WSUS");
    } else {
        // WSUS if we have one, Windows Update if not
        let mut have_wsus = false;
        let service_collection = unsafe {
            update_service_manager.Services()
        }
            .expect("failed to obtain existing update services");
        let service_count = unsafe {
            service_collection.Count()
        }
            .expect("failed to obtain number of existing update services");
        for i in 0..service_count {
            let service = unsafe {
                service_collection.get_Item(i)
            }
                .expect("failed to obtain update service");
            let is_managed = unsafe {
                service.IsManaged()
            }
                .expect("failed to obtain whether the service is managed");
            if is_managed.as_bool() {
                have_wsus = true;
                break;
            }
        }

        unsafe {
            update_searcher.SetServerSelection(if have_wsus { ssManagedServer } else { ssWindowsUpdate })
        }
            .expect("failed to set service selection to WSUS/WU");
    }

    // search
    let criteria = if let Some(c) = &opts.criteria {
        BSTR::from(c)
    } else {
        BSTR::from("IsInstalled = 0 AND Type = 'Software' AND IsHidden = 0")
    };
    println!("searching...");
    let search_results = unsafe {
        update_searcher.Search(
            &criteria,
        )
    }
        .expect("failed to perform search");

    let result_code = unsafe {
        search_results.ResultCode()
    }
        .expect("failed to obtain search operation result code");
    if result_code != orcSucceeded {
        let result_string = result_code_string(result_code);
        println!("search operation result: {}", result_string);
        if result_code != orcSucceededWithErrors {
            // give up immediately
            return ExitCode::FAILURE;
        } else {
            // warn the user once we're through
            exit_code = ExitCode::FAILURE;
        }
    }

    let found_updates_com = unsafe {
        search_results.Updates()
    }
        .expect("failed to obtain found updates");
    let found_updates_count = unsafe {
        found_updates_com.Count()
    }
        .expect("failed to obtain number of found updates");
    let found_updates_len: usize = found_updates_count
        .try_into().expect("failed to convert number of found updates");
    let mut found_updates = Vec::with_capacity(found_updates_len);
    for i in 0..found_updates_count {
        let this_update = unsafe {
            found_updates_com.get_Item(i)
        }
            .expect("failed to obtain update from search results");
        found_updates.push(this_update);
    }

    if found_updates.len() == 0 {
        println!("No updates found.");
        return ExitCode::SUCCESS;
    }

    let mut chosen_updates = Vec::new();
    if opts.interactive {
        for (i, update) in found_updates.iter().enumerate() {
            let name: String = unsafe {
                update.Title()
            }
                .expect("failed to obtain update name")
                .try_into().expect("update name is invalid UTF-16");
            let behavior = unsafe {
                update.InstallationBehavior()
            }
                .expect("failed to obtain installation behavior of update");
            let impact = unsafe {
                behavior.Impact()
            }
                .expect("failed to obtain installation impact of update");
            let reboot = unsafe {
                behavior.RebootBehavior()
            }
                .expect("failedt to obtain reboot behavior of update");

            println!("{}: {}{}{}", i, name, minor_tag(impact), reboot_tag(reboot));
        }

        let mut input = String::new();
        loop {
            chosen_updates.clear();
            input.clear();

            print!("Which updates would you like? (e.g. 1-3,5,7-12) ");
            std::io::stdout().flush()
                .expect("failed to flush stderr");

            std::io::stdin().read_line(&mut input)
                .expect("failed to read line");

            // get ranges by splitting on commas
            for range_str in input.split(',') {
                // get range limits by splitting once on hyphen
                if let Some((first_str, last_str)) = range_str.split_once('-') {
                    // trim whitespace
                    let first_str_no_ws = first_str.trim();
                    let last_str_no_ws = last_str.trim();

                    let invalid_char =
                        first_str_no_ws.chars().any(|c| c <= '0' || c >= '9')
                        || last_str_no_ws.chars().any(|c| c <= '0' || c >= '9')
                    ;
                    if invalid_char {
                        println!("range {:?} contains an invalid char (expected '0' through '9' and maximum one '-')", range_str);
                        continue;
                    }
                    let Ok(first): Result<usize, _> = first_str_no_ws.parse() else {
                        println!("invalid lower bound {:?}", first_str_no_ws);
                        continue;
                    };
                    let Ok(last): Result<usize, _> = last_str_no_ws.parse() else {
                        println!("invalid upper bound {:?}", last_str_no_ws);
                        continue;
                    };

                    if first > last {
                        println!("first {} must be less than last {}", first, last);
                        continue;
                    }
                    if first >= found_updates.len() || last >= found_updates.len() {
                        println!("range {}-{} not valid, we only have updates 0-{}", first, last, found_updates.len()-1);
                        continue;
                    }

                    for i in first..=last {
                        chosen_updates.push(&found_updates[i]);
                    }
                } else {
                    // no hyphen; assume a singular index
                    let str_no_ws = range_str.trim();
                    if str_no_ws.chars().any(|c| c <= '0' || c >= '9') {
                        println!("index {:?} contains an invalid char (expected '0' through '9' and maximum one '-')", range_str);
                        continue;
                    }
                    let Ok(index): Result<usize, _> = str_no_ws.parse() else {
                        println!("invalid index {:?}", str_no_ws);
                        continue;
                    };
                    if index >= found_updates.len() {
                        println!("index {} not valid, we only have updates 0-{}", index, found_updates.len()-1);
                        continue;
                    }
                    chosen_updates.push(&found_updates[index]);
                }
            }
        }
    } else {
        // let fate decide -- alright

        // take the first exclusive update if there is one
        let mut exclusive_update_opt = None;
        for update in &found_updates {
            let behavior = unsafe {
                update.InstallationBehavior()
            }
                .expect("failed to obtain installation behavior for update");
            let impact = unsafe {
                behavior.Impact()
            }
                .expect("failed to obtain update impact");
            if impact == iiRequiresExclusiveHandling {
                // yup, this is an exclusive one
                exclusive_update_opt = Some(update);
                break;
            }
        }

        if let Some(exclusive_update) = exclusive_update_opt {
            // choose only this one
            chosen_updates.push(exclusive_update);
        } else {
            // throw all of them in
            chosen_updates.extend(found_updates.iter());
        }
    }

    if chosen_updates.len() == 0 {
        println!("No updates chosen, downloading and installing no updates.");
        return ExitCode::SUCCESS;
    }

    // fill up our shopping cart
    let shopping_cart_guid = unsafe {
        CLSIDFromProgID(w!("Microsoft.Update.UpdateColl"))
    }
        .expect("failed to find GUID for update collection class");
    let shopping_cart: IUpdateCollection = unsafe {
        CoCreateInstance(
            &shopping_cart_guid,
            None,
            CLSCTX_ALL,
        )
    }
        .expect("failed to create update collection");
    for chosen_update in chosen_updates {
        unsafe {
            shopping_cart.Add(chosen_update)
        }
            .expect("failed to add chosen update to update collection");
    }

    // start downloading the updates
    let continuation = Arc::new(Semaphore::new(0));
    let download_state = DownloadState {
        continuation: Arc::clone(&continuation),
    }
        .into_outer();
    let downloader = unsafe {
        update_session.CreateUpdateDownloader()
    }
        .expect("failed to create update downloader");
    unsafe {
        downloader.SetUpdates(&shopping_cart)
    }
        .expect("failed to set updates on downloader");
    let download_job = unsafe {
        downloader.BeginDownload(
            download_state.as_interface_ref(),
            download_state.as_interface_ref(),
            &null_variant(),
        )
    }
        .expect("failed to start update download");

    // wait on the semaphore for it to finish
    continuation.decrement_blocking();

    println!("download done");

    let download_result = unsafe {
        downloader.EndDownload(&download_job)
    }
        .expect("failed to obtain download result");

    let result_code = unsafe {
        download_result.ResultCode()
    }
        .expect("failed to obtain download operation result code");
    if result_code != orcSucceeded {
        let result_string = result_code_string(result_code);
        println!("download operation result: {}", result_string);
        if result_code != orcSucceededWithErrors {
            // give up immediately
            return ExitCode::FAILURE;
        } else {
            // warn the user once we're through
            exit_code = ExitCode::FAILURE;
        }
    }

    // whatever was downloaded, install it
    let continuation = Arc::new(Semaphore::new(0));
    let install_state = InstallState {
        continuation: Arc::clone(&continuation),
    }
        .into_outer();
    let installer = unsafe {
        update_session.CreateUpdateInstaller()
    }
        .expect("failed to create update installer");
    unsafe {
        installer.SetUpdates(&shopping_cart)
    }
        .expect("failed to tell installer which updates to install");
    let install_job = unsafe {
        installer.BeginInstall(
            install_state.as_interface_ref(),
            install_state.as_interface_ref(),
            &null_variant(),
        )
    }
        .expect("failed to start installation");

    // wait on the semaphore for it to finish
    continuation.decrement_blocking();

    println!("install done");

    let install_result = unsafe {
        installer.EndInstall(&install_job)
    }
        .expect("failed to obtain installation result");

    let result_code = unsafe {
        install_result.ResultCode()
    }
        .expect("failed to obtain installation operation result code");
    if result_code != orcSucceeded {
        let result_string = result_code_string(result_code);
        println!("installation operation result: {}", result_string);
        if result_code != orcSucceededWithErrors {
            // give up immediately
            return ExitCode::FAILURE;
        } else {
            // warn the user once we're through
            exit_code = ExitCode::FAILURE;
        }
    }

    let gotta_reboot = unsafe {
        install_result.RebootRequired()
    }
        .expect("failed to obtain whether to reboot")
        .as_bool();
    if gotta_reboot {
        if !opts.reboot_when_done {
            println!("gotta reboot the system");
        } else {
            // alright then
            unsafe {
                InitiateSystemShutdownW(
                    PCWSTR(null_mut()), // local machine
                    PCWSTR(null_mut()), // no message
                    0, // no countdown
                    false, // don't force-close apps
                    true, // reboot after shutdown
                )
                    .expect("failed to initiate reboot, you'll have to do it yourself");
            }
        }
    }

    exit_code
}
