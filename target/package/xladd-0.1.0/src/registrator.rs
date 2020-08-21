use variant::Variant;
use entrypoint::excel12;
use xlcall::{ xlGetName, xlfRegister };
use std::ffi::CString;
use winapi::um::debugapi::OutputDebugStringA;

/// Allow xlls to register their exported functions with Excel so they can be
/// used in a spreadsheet or macro. These functions can only be called from 
/// within an implementation of xlAutoOpen.
pub struct Reg {
    dll_name: Variant
}

impl Reg {
    /// Creates a registrator. Internally, it finds the name of this dll.
    pub fn new() -> Reg {
        let dll_name = excel12(xlGetName, &mut []);
        debug_print(&format!("addin loaded from: {}", dll_name));

        Reg { dll_name }
    }

    /// Adds an exported function to Excel. This function can only be called from within
    /// xlAutoOpen.
    /// 
    /// # Arguments
    ///
    /// * `name` - The exported name and also the name that appears in Excel
    /// * `arg_types` - A string describing the return type followed by the arguments.
    /// * `arg_text` - A string showing the arguments in human-readable form
    /// * `category` - Either a built-in category such as Information or your own choice
    /// * `help_text` - A short help description for the function wizard
    /// * `arg_help` - An optional slice of strings showing detailed help for each argument
    /// 
    /// Our recommendation is that the name has some prefix that is unique to your addin,
    /// to prevent clashes with other addins. The arg_types string has a letter for the
    /// return type followed by letters for each argument. The letters are defined in the
    /// Excel SDK, but useful ones include:
    /// 
    /// * `Q` - XLOPER12 Variant argument
    /// * `X` - Pending XLOPER12 for async use
    /// * `A` - Boolean (actually i16 that is zero or one)
    /// * `B` - Double (f64)
    /// * `J` - Integer (i32)
    /// 
    /// The string and array types are geared more for a C or C++ user. My recommendation is
    /// that for these arguments, you accept a Q argument, then use the methods on the
    /// Variant type to unpack them. This may be better for other arguments as well, as you
    /// then have control over the coercion and error handling where the arguments are the
    /// wrong type.
    /// 
    /// The string may be terminated by the following special characters
    /// 
    /// * `!` - Marks the function as volatile, so it is assumed to need calling every calc
    /// * `$` - Marks the function as threadsafe, so it can be called from any thread
    /// * `#` - Allows the function to be called even before the args are evaluated
    ///
    /// # Example
    /// 
    /// reg.add("myAdd", "QQQ$", "first, second", "MyCategory", "Adds two numbers or ranges"
    ///     &["help for first arg", "help for second arg"]);
    ///
    pub fn add(
        &self,
        name: &str,
        arg_types: &str,
        arg_text: &str,
        category: &str,
        help_text: &str,
        arg_help: &[&str]) {

        let mut opers = vec![
            self.dll_name.clone(),
            Variant::from_str(name),
            Variant::from_str(arg_types),
            Variant::from_str(name),
            Variant::from_str(arg_text),
            Variant::from_int(1),            // type 1 means useable anywhere (spreadsheet or macro code)
            Variant::from_str(category),
            Variant::missing(),              // no shortcut
            Variant::missing(),              // no help url for now. If we add it, it needn't mean another argument to add
            Variant::from_str(help_text)];

        // append any argument help strings
        for arg in arg_help.iter() {
            opers.push(Variant::from_str(arg));
        }

        let result = excel12(xlfRegister, opers.as_mut_slice());
        debug_print(&format!("Registered {}: result = {}", name, result));
    }
}

pub fn debug_print(message: &str) {
    let cstr = CString::new(message).unwrap();
    unsafe { OutputDebugStringA(cstr.as_ptr()) };
}
