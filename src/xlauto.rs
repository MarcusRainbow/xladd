//! Functions that are exported from the xll and invoked by Excel
//! The only two essential functions are xlAutoOpen and xlAutoFree12.
//! The first of these is implemented by the dll that uses xladd, as
//! only it knows what it wants to export. Other xlAuto methods can be added
//! here as required.

use xlcall::LPXLOPER12;

#[no_mangle]
pub extern "stdcall" fn xlAutoFree12(px_free: LPXLOPER12) {
    // take ownership of this xloper. Then when our xloper goes
    // out of scope, its drop method will free any resources.
    unsafe { Box::from_raw(px_free) };
}
