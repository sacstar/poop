# pstcheck

`pstcheck.exe` (one folder up) opens a PST with Microsoft's open-source
[outlook-pst](https://crates.io/crates/outlook-pst) library and lists every
problem it finds: folders, folder lists, item lists, items, recipients and
attachments. Drag a PST onto `pstcheck.exe`, or run `pstcheck file.pst`.

It reads only; it never changes the file.

## Building

`pstcheck.rs` is an example program for the outlook-pst crate (version 1.2.0):

1. Download the crate source (`outlook-pst-1.2.0.crate` from crates.io) and unpack it.
2. Copy `pstcheck.rs` into its `examples` folder and add to its `Cargo.toml`:

   ```toml
   [[example]]
   name = "pstcheck"
   path = "examples/pstcheck.rs"
   ```
3. `cargo build --release --example pstcheck` (on Windows), or from Linux:
   `rustup target add x86_64-pc-windows-gnu`, install `gcc-mingw-w64-x86-64`, then
   `CARGO_TARGET_X86_64_PC_WINDOWS_GNU_LINKER=x86_64-w64-mingw32-gcc cargo build --release --example pstcheck --target x86_64-pc-windows-gnu`.
