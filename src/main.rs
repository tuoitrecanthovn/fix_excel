use anyhow::{Context, Result};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use rayon::prelude::*;
use std::fs::{self, File};
use std::io::{BufReader, BufWriter};
use std::path::{Path, PathBuf};
use tempfile::tempdir;
use walkdir::WalkDir;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};



#[derive(Debug, Clone, Copy)]
struct UsedRange {
    last_row: u32,
    last_col: u32,
}

fn local_name(name: &[u8]) -> &str {
    // name có thể dạng "{ns}tag" hoặc "tag"
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rsplit_once('}') {
        Some((_, tag)) => tag,
        None => s,
    }
}

fn col_letters_to_index(s: &str) -> Option<u32> {
    let mut n: u32 = 0;
    for ch in s.chars() {
        if !('A'..='Z').contains(&ch) {
            return None;
        }
        n = n * 26 + (ch as u8 - b'A' + 1) as u32;
    }
    Some(n)
}

fn col_index_to_letters(mut idx: u32) -> String {
    let mut s = Vec::new();
    while idx > 0 {
        let r = (idx - 1) % 26;
        s.push((b'A' + (r as u8)) as char);
        idx = (idx - 1) / 26;
    }
    s.into_iter().rev().collect()
}

fn split_cell_ref(r: &str) -> Option<(u32, u32)> {
    // "BC12" -> (55, 12)
    let pos = r.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = r.split_at(pos);
    let col = col_letters_to_index(letters)?;
    let row: u32 = digits.parse().ok()?;
    Some((col, row))
}

fn find_used_range_sheet(xml_path: &Path) -> Result<UsedRange> {
    let mut reader = Reader::from_file(xml_path)?;
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut last_row: u32 = 0;
    let mut last_col: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if local_name(e.name().as_ref()) == "c" => {
                // lấy r attr
                let mut r_attr: Option<String> = None;
                for a in e.attributes().with_checks(false) {
                    if let Ok(a) = a {
                        if a.key.as_ref() == b"r" {
                            r_attr = Some(String::from_utf8_lossy(&a.value).to_string());
                        }
                    }
                }
                // Đọc đến </c>, kiểm tra có v/f/is
                let mut depth = 1usize;
                let mut seen_v = false;
                let mut seen_f = false;
                let mut seen_is = false;
                let mut inner = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner) {
                        Ok(Event::Start(se)) => {
                            let name = se.name();
                            let tag = local_name(name.as_ref());
                            if tag == "v" {
                                seen_v = true;
                            } else if tag == "f" {
                                seen_f = true;
                            } else if tag == "is" {
                                seen_is = true;
                            }
                            depth += 1;
                        }
                        Ok(Event::Empty(se)) => {
                            let name = se.name();
                            let tag = local_name(name.as_ref());
                            if tag == "v" {
                                seen_v = true;
                            } else if tag == "f" {
                                seen_f = true;
                            } else if tag == "is" {
                                seen_is = true;
                            }
                        }
                        Ok(Event::End(_)) => {
                            depth -= 1;
                            if depth == 0 {
                                break;
                            }
                        }
                        Ok(Event::Eof) => break,
                        Ok(_) => {}
                        Err(e) => return Err(e.into()),
                    }
                    inner.clear();
                }

                if (seen_v || seen_f || seen_is) && r_attr.is_some() {
                    if let Some((c, r)) = split_cell_ref(&r_attr.unwrap()) {
                        if r > last_row {
                            last_row = r;
                        }
                        if c > last_col {
                            last_col = c;
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(e.into()),
        }
        buf.clear();
    }

    Ok(UsedRange { last_row, last_col })
}

/// Pass 2: ghi lại sheet, cắt hàng/cột vượt vùng dùng & dọn các khối phình size
fn rewrite_sheet(xml_in: &Path, xml_out: &Path, used: UsedRange) -> Result<()> {
    let mut reader = Reader::from_file(xml_in)?;
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(BufWriter::new(File::create(xml_out)?));
    let mut buf = Vec::new();

    let mut _current_row_idx: Option<u32> = None;
    let mut drop_this_row = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let tag = local_name(name.as_ref());

                match tag {
                    "dimension" => {
                        // viết lại dimension với ref mới
                        let mut el = BytesStart::new("dimension");
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() != b"ref" {
                                el.push_attribute((
                                    std::str::from_utf8(a.key.as_ref()).unwrap_or(""),
                                    std::str::from_utf8(&a.value).unwrap_or(""),
                                ));
                            }
                        }
                        let new_ref = if used.last_row == 0 || used.last_col == 0 {
                            "A1:A1".to_string()
                        } else {
                            format!("A1:{}{}", col_index_to_letters(used.last_col), used.last_row)
                        };
                        el.push_attribute(("ref", new_ref.as_str()));
                        writer.write_event(Event::Start(el))?;
                    }
                    "row" => {
                        // xác định row index
                        let mut r_idx = None::<u32>;
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() == b"r" {
                                r_idx = std::str::from_utf8(&a.value).ok().and_then(|s| s.parse().ok());
                            }
                        }
                        _current_row_idx = r_idx;
                        drop_this_row = r_idx
                            .map(|r| used.last_row > 0 && r > used.last_row)
                            .unwrap_or(false);

                        if drop_this_row {
                            // ăn hết nội dung <row>…</row> mà không ghi
                            let mut depth = 1usize;
                            let mut inner = Vec::new();
                            loop {
                                match reader.read_event_into(&mut inner) {
                                    Ok(Event::Start(_)) => depth += 1,
                                    Ok(Event::End(_)) => {
                                        depth -= 1;
                                        if depth == 0 {
                                            break;
                                        }
                                    }
                                    Ok(Event::Eof) => break,
                                    Ok(_) => {}
                                    Err(e) => return Err(e.into()),
                                }
                                inner.clear();
                            }
                            _current_row_idx = None;
                            continue;
                        } else {
                            writer.write_event(Event::Start(e.clone()))?;
                        }
                    }
                    "c" => {
                        // kiểm tra cột của cell, nếu > last_col thì bỏ
                        let mut r_attr: Option<String> = None;
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() == b"r" {
                                r_attr =
                                    Some(String::from_utf8_lossy(&a.value).to_string());
                            }
                        }
                        let mut keep = true;
                        if let Some(r) = r_attr.as_ref() {
                            if let Some((c, _)) = split_cell_ref(r) {
                                if used.last_col > 0 && c > used.last_col {
                                    keep = false;
                                }
                            }
                        }
                        if !keep {
                            // skip cả block <c>…</c>
                            let mut depth = 1usize;
                            let mut inner = Vec::new();
                            loop {
                                match reader.read_event_into(&mut inner) {
                                    Ok(Event::Start(_)) => depth += 1,
                                    Ok(Event::End(_)) => {
                                        depth -= 1;
                                        if depth == 0 {
                                            break;
                                        }
                                    }
                                    Ok(Event::Eof) => break,
                                    Ok(_) => {}
                                    Err(e) => return Err(e.into()),
                                }
                                inner.clear();
                            }
                            continue;
                        } else {
                            writer.write_event(Event::Start(e.clone()))?;
                        }
                    }
                    "mergeCells" => {
                        // bắt & lọc toàn bộ mergeCells rồi viết lại
                        let mut kept: Vec<(u32, u32, u32, u32)> = Vec::new();
                        let mut inner = Vec::new();
                        loop {
                            match reader.read_event_into(&mut inner) {
                                Ok(Event::Empty(ref mc)) => {
                                    if local_name(mc.name().as_ref()) == "mergeCell" {
                                        for a in mc.attributes().with_checks(false).flatten() {
                                            if a.key.as_ref() == b"ref" {
                                                if let Ok(s) = std::str::from_utf8(&a.value) {
                                                    if let Some((a, b)) = s.split_once(':') {
                                                        if let (Some((c1, r1)), Some((c2, r2))) =
                                                            (split_cell_ref(a), split_cell_ref(b))
                                                        {
                                                            if c1 >= 1
                                                                && c2 >= 1
                                                                && r1 >= 1
                                                                && r2 >= 1
                                                                && (used.last_col == 0
                                                                    || (c1 <= used.last_col
                                                                        && c2 <= used.last_col))
                                                                && (used.last_row == 0
                                                                    || (r1 <= used.last_row
                                                                        && r2 <= used.last_row))
                                                            {
                                                                kept.push((c1, r1, c2, r2));
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                Ok(Event::Start(ref mc)) => {
                                    if local_name(mc.name().as_ref()) == "mergeCell" {
                                        // đọc attr rồi nhảy đến </mergeCell>
                                        let mut ref_str: Option<String> = None;
                                        for a in mc.attributes().with_checks(false).flatten() {
                                            if a.key.as_ref() == b"ref" {
                                                ref_str = Some(
                                                    String::from_utf8_lossy(&a.value).to_string(),
                                                );
                                            }
                                        }
                                        if let Some(s) = ref_str {
                                            if let Some((a, b)) = s.split_once(':') {
                                                if let (Some((c1, r1)), Some((c2, r2))) =
                                                    (split_cell_ref(a), split_cell_ref(b))
                                                {
                                                    if c1 >= 1
                                                        && c2 >= 1
                                                        && r1 >= 1
                                                        && r2 >= 1
                                                        && (used.last_col == 0
                                                            || (c1 <= used.last_col
                                                                && c2 <= used.last_col))
                                                        && (used.last_row == 0
                                                            || (r1 <= used.last_row
                                                                && r2 <= used.last_row))
                                                    {
                                                        kept.push((c1, r1, c2, r2));
                                                    }
                                                }
                                            }
                                        }
                                        reader.read_to_end_into(mc.name(), &mut Vec::new())?;
                                    } else {
                                        // bỏ qua phần tử con khác
                                        reader.read_to_end_into(mc.name(), &mut Vec::new())?;
                                    }
                                }
                                Ok(Event::End(ref ee)) => {
                                    if local_name(ee.name().as_ref()) == "mergeCells" {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                Ok(_) => {}
                                Err(e) => return Err(e.into()),
                            }
                            inner.clear();
                        }

                        // ghi lại mergeCells nếu còn
                        if !kept.is_empty() {
                            let mut mc_s = BytesStart::new("mergeCells");
                            mc_s.push_attribute(("count", kept.len().to_string().as_str()));
                            writer.write_event(Event::Start(mc_s))?;
                            for (c1, r1, c2, r2) in kept.drain(..) {
                                let mut m = BytesStart::new("mergeCell");
                                let r = format!(
                                    "{}{}:{}{}",
                                    col_index_to_letters(c1),
                                    r1,
                                    col_index_to_letters(c2),
                                    r2
                                );
                                m.push_attribute(("ref", r.as_str()));
                                writer.write_event(Event::Empty(m))?;
                            }
                            writer.write_event(Event::End(BytesEnd::new("mergeCells")))?;
                        }
                    }
                    // Dọn các khối "nặng": skip toàn bộ
                    "conditionalFormatting"
                    | "dataValidations"
                    | "pageBreaks"
                    | "ignoredErrors"
                    | "extLst"
                    | "phoneticPr" => {
                        reader.read_to_end_into(e.name(), &mut Vec::new())?;
                        // không ghi gì
                    }
                    _ => {
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let tag = local_name(name.as_ref());
                match tag {
                    "dimension" => {
                        let mut el = BytesStart::new("dimension");
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() != b"ref" {
                                el.push_attribute((
                                    std::str::from_utf8(a.key.as_ref()).unwrap_or(""),
                                    std::str::from_utf8(&a.value).unwrap_or(""),
                                ));
                            }
                        }
                        let new_ref = if used.last_row == 0 || used.last_col == 0 {
                            "A1:A1".to_string()
                        } else {
                            format!("A1:{}{}", col_index_to_letters(used.last_col), used.last_row)
                        };
                        el.push_attribute(("ref", new_ref.as_str()));
                        writer.write_event(Event::Empty(el))?;
                    }
                    // skip các singleton nặng nếu có
                    "conditionalFormatting"
                    | "dataValidations"
                    | "pageBreaks"
                    | "ignoredErrors"
                    | "extLst"
                    | "phoneticPr" => {
                        // không ghi gì
                    }
                    _ => writer.write_event(Event::Empty(e.clone()))?,
                }
            }
            Ok(Event::End(e)) => {
                writer.write_event(Event::End(e))?;
            }
            Ok(Event::Text(t)) => {
                writer.write_event(Event::Text(t))?;
            }
            Ok(Event::Decl(d)) => writer.write_event(Event::Decl(d))?,
            Ok(Event::PI(p)) => writer.write_event(Event::PI(p))?,
            Ok(Event::CData(c)) => writer.write_event(Event::CData(c))?,
            Ok(Event::Comment(_)) => { /* drop comments */ }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.into()),
            _ => (), // Ignore other events
        }
        buf.clear();
    }

    Ok(())
}

fn trim_one_xlsx(input: &Path, output: &Path) -> Result<()> {
    // 1) extract zip vào thư mục tạm
    let tmp = tempdir()?;
    let tmpdir = tmp.path();

    // unzip
    {
        let f = File::open(input)?;
        let mut zin = ZipArchive::new(f)?;
        for i in 0..zin.len() {
            let mut file = zin.by_index(i)?;
            let out_path = tmpdir.join(file.name());
            if file.is_dir() {
                fs::create_dir_all(&out_path)?;
            } else {
                if let Some(parent) = out_path.parent() {
                    fs::create_dir_all(parent)?;
                }
                let mut out = BufWriter::new(File::create(&out_path)?);
                std::io::copy(&mut file, &mut out)?;
            }
        }
    }

    // 2) xử lý xl/worksheets/*.xml song song
    let ws_dir = tmpdir.join("xl/worksheets");
    if ws_dir.exists() {
        let sheets: Vec<PathBuf> = fs::read_dir(&ws_dir)?
            .filter_map(|e| e.ok().map(|e| e.path()))
            .filter(|p| p.extension().map(|x| x == "xml").unwrap_or(false))
            .collect();

        sheets.par_iter().try_for_each(|sheet_xml| -> Result<()> {
            let used = find_used_range_sheet(sheet_xml)
                .with_context(|| format!("find_used_range {}", sheet_xml.display()))?;
            let tmp_out = sheet_xml.with_extension("xml.out");
            rewrite_sheet(sheet_xml, &tmp_out, used)
                .with_context(|| format!("rewrite_sheet {}", sheet_xml.display()))?;
            fs::rename(&tmp_out, sheet_xml)?;
            Ok(())
        })?;
    }

    // 3) xoá calcChain.xml (Excel tự rebuild)
    let _ = fs::remove_file(tmpdir.join("xl/calcChain.xml"));

    // 4) re-zip
    {
        let f = File::create(output)?;
        let mut zw = ZipWriter::new(BufWriter::new(f));
        let options: FileOptions<()> = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for entry in WalkDir::new(tmpdir).into_iter().filter_map(|e| e.ok()) {
            let path = entry.path();
            if path == tmpdir {
                continue;
            }
            let name = path
                .strip_prefix(tmpdir)
                .unwrap()
                .to_string_lossy()
                .replace('\\', "/");
            if path.is_dir() {
                zw.add_directory(name, options)?;
            } else {
                zw.start_file(name, options)?;
                let mut f = BufReader::new(File::open(path)?);
                std::io::copy(&mut f, &mut zw)?;
            }
        }
        zw.finish()?;
    }

    Ok(())
}

fn process_path(input: &Path, out_dir: Option<&Path>, threshold_mb: u64, suffix: &str) -> Result<()> {
    let mut files: Vec<PathBuf> = Vec::new();
    if input.is_file() && input.extension().map(|e| e == "xlsx").unwrap_or(false) {
        files.push(input.to_path_buf());
    } else if input.is_dir() {
        for e in WalkDir::new(input).into_iter().filter_map(|e| e.ok()) {
            let p = e.path();
            if p.is_file() && p.extension().map(|e| e == "xlsx").unwrap_or(false) {
                files.push(p.to_path_buf());
            }
        }
    } else {
        anyhow::bail!("Đường dẫn không hợp lệ");
    }

    if let Some(od) = out_dir {
        fs::create_dir_all(od)?;
    }

    for p in files {
        let sz = fs::metadata(&p)?.len();
        let sz_mb = sz / (1024 * 1024);
        if sz_mb < threshold_mb {
            eprintln!("Bỏ qua {} ({} MB <= {} MB)", p.display(), sz_mb, threshold_mb);
            continue;
        }
        let out = if let Some(od) = out_dir {
            od.join(format!(
                "{}{}{}",
                p.file_stem().unwrap().to_string_lossy(),
                suffix,
                ".xlsx"
            ))
        } else {
            p.with_file_name(format!(
                "{}{}{}",
                p.file_stem().unwrap().to_string_lossy(),
                suffix,
                ".xlsx"
            ))
        };
        eprintln!("▶ Xử lý: {} ({} MB) → {}", p.display(), sz_mb, out.display());
        trim_one_xlsx(&p, &out)?;
        let new_sz = fs::metadata(&out)?.len() / (1024 * 1024);
        eprintln!("   ✓ Mới: {} MB (giảm {} MB)", new_sz, (sz_mb as i64 - new_sz as i64));
    }

    Ok(())
}

fn main() -> Result<()> {
    use std::env;
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!(
            "Cách dùng:
  xlsx-trimmer <đường-dẫn-file-hoặc-thư-mục>
    [-o <output-dir>] [--threshold-mb 10] [--suffix _trimmed]"
        );
        std::process::exit(1);
    }
    let mut input = None::<PathBuf>;
    let mut out_dir = None::<PathBuf>;
    let mut threshold: u64 = 10;
    let mut suffix = String::from("_trimmed");

    let mut i = 1;
    while i < args.len() {
        let arg = &args[i];
        match arg.as_str() {
            "-o" | "--output-dir" => {
                i += 1;
                if i >= args.len() {
                    anyhow::bail!("Thiếu giá trị cho tham số '{}'", arg);
                }
                out_dir = Some(PathBuf::from(&args[i]));
            }
            "--threshold-mb" => {
                i += 1;
                if i >= args.len() {
                    anyhow::bail!("Thiếu giá trị cho tham số '{}'", arg);
                }
                threshold = args[i].parse().with_context(|| {
                    format!("Giá trị không hợp lệ cho --threshold-mb: '{}'", args[i])
                })?;
            }
            "--suffix" => {
                i += 1;
                if i >= args.len() {
                    anyhow::bail!("Thiếu giá trị cho tham số '{}'", arg);
                }
                suffix = args[i].clone();
            }
            _ if input.is_none() && !arg.starts_with('-') => {
                input = Some(PathBuf::from(arg));
            }
            other => {
                anyhow::bail!("Tham số không hợp lệ: {}", other);
            }
        }
        i += 1;
    }

    if let Some(input_path) = input {
        process_path(&input_path, out_dir.as_deref(), threshold, &suffix)?;
    } else {
        anyhow::bail!("Thiếu đường dẫn file hoặc thư mục đầu vào.");
    }
    Ok(())
}
