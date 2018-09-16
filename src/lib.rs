pub mod xlcall;
pub mod entrypoint;
pub mod variant;

extern crate winapi;
extern crate widestring;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
