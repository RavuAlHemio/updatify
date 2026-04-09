use std::sync::{Condvar, Mutex, TryLockError};


/// The classical semaphore as described by Edsger W. Dijkstra.
///
/// A counter that offers the following operations:
///
/// * decrement_blocking: Attempts to decrease the value of the semaphore by 1. If the value is
///   already 0, blocks until the value is at least 1 (because another thread called increment),
///   then decreases it by 1 and returns. Dijkstra called this operation P.
///
/// * increment: Increases the value of the semaphore by 1 and, if any threads are blocking within
///   the semaphore's decrement_blocking operation, wakes one of them. Dijkstra named this operation
///   V.
///
/// A semaphore must ensure internally that the operations, apart from the blocking section within
/// the decrement_blocking operation, are completed atomically, i.e. only one thread may ever modify
/// the value of the semaphore at a time.
///
/// This sempahore implementation defers to [`Mutex`] (for the atomicity) and [`Condvar`] (for the
/// conditional blocking and waking) from the standard library. Thus, the implementation of
/// [`Condvar`] ultimately decides which thread is awoken (and thereby whether the load is
/// distributed fairly or it's always the same thread that is allowed to continue).
pub struct Semaphore {
    counter: Mutex<usize>,
    condition: Condvar,
}
impl Semaphore {
    pub fn new(initial_value: usize) -> Self {
        Semaphore {
            counter: Mutex::new(initial_value),
            condition: Condvar::new(),
        }
    }

    /// Attempts to decrement the semaphore's value and blocks until that is possible.
    pub fn decrement_blocking(&self) {
        let mut counter_guard = self.counter
            .lock().expect("poisoned?!");
        loop {
            if *counter_guard > 0 {
                // no need to block
                *counter_guard -= 1;
                break;
            }

            // wait until somebody comes and wakes us
            counter_guard = self.condition.wait(counter_guard)
                .expect("poisoned?!");
        }
    }

    /// Increments the semaphore's value. If there is a thread blocking on decrement_block, wakes
    /// it.
    pub fn increment(&self) {
        {
            let mut counter_guard = self.counter
                .lock().expect("poisoned?!");
            *counter_guard += 1;
        }

        self.condition.notify_one();
    }

    /// Attempts to decrement the semaphore's value and returns whether it was possible. Does not
    /// block.
    ///
    /// This is an extension to Dijkstra's operations.
    pub fn try_decrement(&self) -> bool {
        let mut counter_guard = match self.counter.try_lock() {
            Ok(cg) => cg,
            Err(TryLockError::Poisoned(e)) => {
                panic!("mutex within semaphore is poisoned: {}", e);
            },
            Err(TryLockError::WouldBlock) => return false,
        };
        if *counter_guard > 0 {
            *counter_guard -= 1;
            true
        } else {
            false
        }
    }
}
