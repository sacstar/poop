//! pstcheck: opens a PST with Microsoft's outlook-pst library and reports every problem it finds.
use outlook_pst::{ndb::{header::Header, root::Root}, ltp::table_context::TableContext, messaging::{folder::Folder, store::Store}, ndb::node_id::NodeId, *};
use std::rc::Rc;

struct Report { folders: usize, messages: usize, errors: Vec<String> }

impl Report {
    fn err(&mut self, what: String) {
        if self.errors.len() < 200 { println!("  PROBLEM: {what}"); }
        self.errors.push(what);
    }
}

fn read_table(t: &dyn TableContext, what: &str, r: &mut Report) {
    let ctx = t.context();
    for row in t.rows_matrix() {
        match row.columns(ctx) {
            Ok(cols) => {
                for (col, value) in ctx.columns().iter().zip(cols) {
                    if let Some(v) = value {
                        if let Err(e) = t.read_column(&v, col.prop_type()) {
                            r.err(format!("{what}: row {:X} column {:04X}: {e:?}", u32::from(row.id()), col.prop_id()));
                        }
                    }
                }
            }
            Err(e) => r.err(format!("{what}: row {:X}: {e:?}", u32::from(row.id()))),
        }
    }
}

fn walk(store: &Rc<dyn Store>, folder: Rc<dyn Folder>, path: &str, r: &mut Report) {
    r.folders += 1;
    let mut items = 0;
    match folder.contents_table() {
        Some(ct) => {
            read_table(ct.as_ref(), &format!("item list of {path}"), r);
            for row in ct.rows_matrix() {
                items += 1;
                let id = u32::from(row.id());
                let eid = match store.properties().make_entry_id(NodeId::from(id)) { Ok(e) => e, Err(e) => { r.err(format!("{path}: item {id:X}: {e:?}")); continue; } };
                match store.open_message(&eid, None) {
                    Ok(m) => {
                        r.messages += 1;
                        if let Some(t) = m.recipient_table() { read_table(t.as_ref(), &format!("{path}: recipients of item {id:X}"), r); }
                        if let Some(t) = m.attachment_table() { read_table(t.as_ref(), &format!("{path}: attachments of item {id:X}"), r); }
                    }
                    Err(e) => r.err(format!("{path}: item {id:X} won't open: {e:?}")),
                }
            }
        }
        None => {}
    }
    if let Some(at) = folder.associated_table() { read_table(at.as_ref(), &format!("hidden items of {path}"), r); }
    let count = folder.properties().content_count().unwrap_or(-1);
    println!("{path}  ({items} items, count says {count})");
    if count >= 0 && count as usize != items { r.err(format!("{path}: folder count {count} but {items} items listed")); }
    if let Some(ht) = folder.hierarchy_table() {
        read_table(ht.as_ref(), &format!("subfolder list of {path}"), r);
        for row in ht.rows_matrix() {
            let id = u32::from(row.id());
            match store.properties().make_entry_id(NodeId::from(id)).and_then(|e| store.open_folder(&e)) {
                Ok(sub) => {
                    let name = sub.properties().display_name().unwrap_or_else(|e| { format!("<no name: {e:?}>") });
                    walk(store, sub, &format!("{path}\\{name}"), r);
                }
                Err(e) => r.err(format!("{path}: subfolder {id:X} won't open: {e:?}")),
            }
        }
    }
}

fn main() {
    let Some(path) = std::env::args().nth(1) else {
        eprintln!("Usage: pstcheck <file.pst>   (or drag a PST onto pstcheck.exe)");
        wait();
        return;
    };
    println!("Checking {path}\n");
    let mut r = Report { folders: 0, messages: 0, errors: vec![] };
    match UnicodePstFile::open(&path) {
        Ok(pst) => {
            let h = pst.header();
            println!("Header OK. Encryption: {:?}. AMap status: {:?}", h.crypt_method(), h.root().amap_is_valid());
        }
        Err(e) => r.err(format!("file won't open: {e:?}")),
    }
    match open_store(&path) {
        Ok(store) => {
            match store.properties().display_name() { Ok(n) => println!("Store name: {n}"), Err(e) => r.err(format!("store name: {e:?}")) }
            if let Err(e) = store.named_property_map() { r.err(format!("named property map: {e:?}")); }
            for (what, id) in [("top of folders", store.properties().ipm_sub_tree_entry_id()), ("deleted items", store.properties().ipm_wastebasket_entry_id()), ("search root", store.properties().finder_entry_id())] {
                match id.and_then(|e| store.open_folder(&e)) { Ok(_) => {}, Err(e) => r.err(format!("{what} folder: {e:?}")) }
            }
            match store.properties().make_entry_id(NodeId::from(0x122_u32)).and_then(|e| store.open_folder(&e)) {
                Ok(root) => walk(&store, root, "(root)", &mut r),
                Err(e) => r.err(format!("root folder: {e:?}")),
            }
        }
        Err(e) => r.err(format!("store won't open: {e:?}")),
    }
    println!("\nChecked {} folders, opened {} items. Problems: {}", r.folders, r.messages, r.errors.len());
    if r.errors.is_empty() { println!("RESULT: OK"); } else { println!("RESULT: {} PROBLEM(S) - first: {}", r.errors.len(), r.errors[0]); }
    wait();
}

fn wait() {
    if std::env::var("PSTCHECK_NOWAIT").is_err() {
        println!("\nPress Enter to close.");
        let mut s = String::new();
        let _ = std::io::stdin().read_line(&mut s);
    }
}
