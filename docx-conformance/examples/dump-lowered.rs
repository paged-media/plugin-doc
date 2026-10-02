//! Dump a .docx's lowering as JSON (styles, blocks, diagnostics): the
//! debugging companion of scripts/real-docx-acceptance.sh.
//!   cargo run -p docx-conformance --example dump-lowered -- <file.docx> > out.json
fn main() {
    let path = std::env::args()
        .nth(1)
        .expect("usage: dump-lowered <file.docx>");
    let bytes = std::fs::read(&path).expect("read docx");
    let doc = docx_import::import_docx(&bytes).expect("import");
    let lowered = docx_lower::lower(&doc);
    println!("{}", serde_json::to_string(&lowered).expect("json"));
}
