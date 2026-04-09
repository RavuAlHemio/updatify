use std::os::windows::ffi::OsStrExt;
use std::path::PathBuf;

use clap::Parser;
use windows::core::{BSTR, Error, Ref, Interface, implement, w};
use windows::Win32::System::Com::{
    CLSCTX_ALL, CLSIDFromProgID, COINIT_MULTITHREADED, CoCreateInstance, CoInitializeEx,
};
use windows::Win32::System::UpdateAgent::{
    ISearchCompletedCallback, ISearchCompletedCallbackArgs, ISearchCompletedCallback_Impl,
    ISearchJob, IUpdateServiceManager, IUpdateServiceManager2, IUpdateSession,
    asfAllowOnlineRegistration, asfAllowPendingRegistration, asfRegisterServiceWithAU,
    ssManagedServer, ssOthers, ssWindowsUpdate,
};


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
    pub criteria: Option<String>,
}

#[implement(ISearchCompletedCallback)]
struct DoneSearching;
impl ISearchCompletedCallback_Impl for DoneSearching_Impl {
    fn Invoke(
        &self,
        search_job: Ref<ISearchJob>,
        callback_args: Ref<ISearchCompletedCallbackArgs>,
    ) -> Result<(), Error> {
        todo!()
    }
}


fn main() {
    let opts = Opts::parse();

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

    let mut offline_service_opt = None;
    if opts.local.is_some() || opts.ms_update_opt_in {
        // we will need an update service manager
        let update_service_manager: IUpdateServiceManager = unsafe {
            CoCreateInstance(
                &update_service_manager_guid,
                None,
                CLSCTX_ALL,
            )
        }
            .expect("failed to create update service manager");

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
    if let Some(offline_service) = offline_service_opt {
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
        unsafe {
            update_searcher.SetServerSelection(ssManagedServer)
        }
            .expect("failed to set service selection to WSUS");
    }

    // search
    let criteria = if let Some(c) = &opts.criteria {
        BSTR::from(c)
    } else {
        BSTR::from("IsInstalled = 0 AND Type = 'Software' AND IsHidden = 0")
    };

    todo!();
    /*
    update_searcher.BeginSearch(
        &criteria,
        oncompleted,
        state,
    )
    */
}
