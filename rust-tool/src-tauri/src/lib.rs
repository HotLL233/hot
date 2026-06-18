use calamine::{DataType, Reader, Xlsx};
use chrono::NaiveDate;
use regex::Regex;
use rust_xlsxwriter::*;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::path::Path;
use std::sync::Mutex;
use tauri::State;
use std::fs;

// ===================== 数据结构 =====================

#[derive(Debug, Serialize, Deserialize, Clone)]
struct FileResult {
    files: Vec<String>,
    results: Option<TableData>,
    status: String,
    imported_count: usize,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
struct TableData {
    columns: Vec<String>,
    rows: Vec<Vec<String>>,
    totals: Vec<String>,
    mode: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
struct UpdateInfo {
    has_update: bool,
    latest_version: String,
    current_version: String,
    release_notes: String,
    download_url: String,
}

// ===================== 应用状态 =====================

struct AppState {
    file_paths: Mutex<Vec<String>>,
    merged_results: Mutex<Option<TableData>>,
    export_dir: Mutex<String>,
    logs: Mutex<Vec<String>>,
}

const VERSION: &str = "5.0.0";
const DATE_COL: &str = "日期";
const BATCH_COL: &str = "批号";
const SKIP_ROWS: usize = 2;
const DATE_KEYWORDS: &[&str] = &[
    "日期","date","时间","time","检测日期","采样日期",
    "报告日期","分析日期","送检日期","检验日期",
];
const BATCH_KEYWORDS: &[&str] = &[
    "批号","batch","lot","编号","序号",
    "样品批号","样品编号","样品号","样品名称",
    "实验编号","样本号","code","id",
];

// ===================== 日志 =====================

fn app_log(state: &AppState, msg: String) {
    if let Ok(mut logs) = state.logs.lock() {
        let ts = chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
        logs.push(format!("[{}] {}", ts, msg));
        if logs.len() > 500 { logs.remove(0); }
    }
}

// ===================== 日期/数值处理 =====================

fn year_ok(d: &NaiveDate) -> bool {
    let y: i32 = d.format("%Y").to_string().parse().unwrap_or(0);
    y >= 1900 && y <= 2100
}

fn try_parse_yyyymmdd(s: &str) -> Option<String> {
    let clean = if let Some(dot) = s.find('.') { &s[..dot] } else { s };
    if clean.len() == 8 && clean.chars().all(|c| c.is_ascii_digit()) {
        if let Ok(d) = NaiveDate::parse_from_str(clean, "%Y%m%d") {
            if year_ok(&d) { return Some(d.format("%Y-%m-%d").to_string()); }
        }
    }
    None
}

fn ensure_date_str(v: &dyn DataType) -> String {
    // Try as string
    if let Some(s) = v.as_string() {
        let s = s.trim().to_string();
        if s.is_empty() { return String::new(); }
        // YYYYMMDD string
        if let Some(d) = try_parse_yyyymmdd(&s) { return d; }
        // Common formats
        for fmt in &["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"] {
            if let Ok(d) = NaiveDate::parse_from_str(&s, fmt) { return d.format("%Y-%m-%d").to_string(); }
        }
        return String::new();
    }
    // Try as float
    if let Some(f) = v.as_f64() {
        let s = format!("{:.0}", f);
        // YYYYMMDD integer → "20250101"
        if let Some(d) = try_parse_yyyymmdd(&s) { return d; }
        // Excel serial date
        if f >= 40000.0 && f <= 200000.0 {
            if let Some(d) = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
                .checked_add_signed(chrono::Duration::days(f as i64)) {
                if year_ok(&d) { return d.format("%Y-%m-%d").to_string(); }
            }
        }
        return String::new();
    }
    String::new()
}

fn get_as_string(v: &dyn DataType) -> String {
    if let Some(s) = v.as_string() { return s.trim().to_string(); }
    if let Some(f) = v.as_f64() {
        if f.fract() == 0.0 && f >= 0.0 { return format!("{:.0}", f); }
        return f.to_string();
    }
    if let Some(i) = v.as_i64() { return i.to_string(); }
    String::new()
}

// ===================== 检测逻辑（与原版一致） =====================

fn detect_data_start_row(range: &calamine::Range<calamine::Data>, max_scan: usize) -> usize {
    let re = Regex::new(r"^\d{4}-\d{2}-\d{2}$").unwrap();
    let ncols = (range.width() as usize).min(30);
    let nrows = (range.height() as usize).min(max_scan);
    for row_idx in 0..nrows {
        for col_idx in 0..ncols {
            if let Some(val) = range.get((row_idx, col_idx)) {
                if re.is_match(&ensure_date_str(val)) { return row_idx; }
            }
        }
    }
    SKIP_ROWS
}

fn detect_columns(range: &calamine::Range<calamine::Data>, header_end: usize) -> (usize, usize) {
    let max_header_scan: usize = 3;
    let mut date_col = 0usize; let mut batch_col = 1usize;
    let mut df = false; let mut bf = false;
    let h = header_end.min(range.height() as usize);
    let w = (range.width() as usize).min(30);
    if h == 0 || w < 2 { return (date_col, batch_col); }
    let scan_start = if h > max_header_scan { h - max_header_scan } else { 0 };
    for row_idx in (scan_start..h).rev() {
        for col_idx in 0..w {
            if let Some(val) = range.get((row_idx, col_idx)) {
                let s = val.as_string().unwrap_or_default().to_lowercase();
                if s.trim().is_empty() { continue; }
                if !df && DATE_KEYWORDS.iter().any(|kw| s.contains(kw)) { date_col = col_idx; df = true; }
                if !bf && BATCH_KEYWORDS.iter().any(|kw| s.contains(kw)) { batch_col = col_idx; bf = true; }
            }
        }
        if df && bf { break; }
    }
    (date_col, batch_col)
}

fn process_sheet(range: &calamine::Range<calamine::Data>) -> Option<Vec<(String, String)>> {
    let auto_skip = detect_data_start_row(range, 50);
    let h = range.height() as usize;
    let w = range.width() as usize;
    
    // 列检测：只扫描表头区域（前 auto_skip 行，从最后 3 行向上）
    let header_rows = auto_skip.min(h);
    let (dc, bc) = if header_rows > 0 {
        detect_columns(range, header_rows)
    } else {
        (0usize, 1usize)
    };
    if dc >= w || bc >= w { return None; }
    
    let re = Regex::new(r"^\d{4}-\d{2}-\d{2}$").unwrap();
    let mut records: Vec<(String, String)> = Vec::new();
    let mut last_date = String::new();
    let prefixes = ["例","示例","sample"];
    
    for row_idx in auto_skip..h {
        let dv = range.get((row_idx, dc)).map(|v| ensure_date_str(v)).unwrap_or_default();
        let date = if dv.is_empty() { last_date.clone() } else { last_date = dv.clone(); dv };
        if !re.is_match(&date) { continue; }
        
        let bv = range.get((row_idx, bc)).map(|v| get_as_string(v)).unwrap_or_default();
        if bv.is_empty() { continue; }
        if prefixes.iter().any(|p| bv.starts_with(p)) { continue; }
        records.push((date, bv));
    }
    if records.is_empty() { None } else { Some(records) }
}

// ===================== 统计算法 =====================

/// 按 (sheet_name) 聚合各文件统计，相同 sheet 名跨文件累加（与原版一致）
fn process_all_files(paths: &[String]) -> (Vec<String>, TableData) {
    let mut sheet_counts: BTreeMap<String, BTreeMap<String, usize>> = BTreeMap::new();
    let mut sheet_order: Vec<String> = Vec::new();
    
    for path in paths {
        let mut wb = match calamine::open_workbook::<Xlsx<_>, _>(path.clone()) { Ok(wb) => wb, Err(_) => continue, };
        for name in wb.sheet_names().to_vec() {
            if name == "目录" { continue; }
            let range = match wb.worksheet_range(&name) { Ok(r) => r, Err(_) => continue, };
            let records = match process_sheet(&range) { Some(r) => r, None => continue, };
            
            // 按日期统计本 sheet 的批号去重数
            let mut daily: BTreeMap<String, HashSet<String>> = BTreeMap::new();
            for (d, b) in &records { daily.entry(d.clone()).or_default().insert(b.clone()); }
            
            // 聚合到 sheet 名下（同名跨文件累加）
            let entry = sheet_counts.entry(name.clone()).or_default();
            for (date, batches) in &daily {
                let cnt = entry.entry(date.clone()).or_insert(0);
                *cnt += batches.len();
            }
            
            if !sheet_order.contains(&name) { sheet_order.push(name); }
        }
    }
    
    // 构建宽表（与原版一致：columns=[日期, sheet1, sheet2, ..., 合计]）
    let mut columns = vec![DATE_COL.to_string()];
    columns.extend(sheet_order.clone());
    columns.push("合计".to_string());
    
    let mut all_dates: BTreeSet<String> = BTreeSet::new();
    for sheets in sheet_counts.values() { for d in sheets.keys() { all_dates.insert(d.clone()); } }
    
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut tot: Vec<usize> = vec![0; sheet_order.len()];
    let mut grand_total = 0usize;
    
    for date in &all_dates {
        let mut row: Vec<String> = vec![date.clone()];
        let mut row_total = 0usize;
        for (si, sheet) in sheet_order.iter().enumerate() {
            let val = sheet_counts.get(sheet).and_then(|m| m.get(date.as_str())).copied().unwrap_or(0);
            row.push(if val == 0 { String::new() } else { val.to_string() });
            row_total += val;
            tot[si] += val;
        }
        row.push(if row_total == 0 { String::new() } else { row_total.to_string() });
        grand_total += row_total;
        rows.push(row);
    }
    
    let mut totals: Vec<String> = vec!["合计".to_string()];
    for t in &tot { totals.push(if *t == 0 { String::new() } else { t.to_string() }); }
    totals.push(if grand_total == 0 { String::new() } else { grand_total.to_string() });
    
    rows.push(totals.clone());
    
    let files: Vec<String> = paths.iter().map(|p| Path::new(p).file_name().unwrap().to_string_lossy().to_string()).collect();
    (files, TableData { columns, rows, totals, mode: "全部文件".to_string() })
}

/// 构建宽表（按 Sheet 分列展示）：日期 | SheetA | SheetB | ... | 合计
fn build_wide_table(per_sheet: &[(String, BTreeMap<String, BTreeSet<String>>)], mode: &str) -> TableData {
    let sheet_names: Vec<&String> = per_sheet.iter().map(|(name, _)| name).collect();
    let mut all_dates: BTreeSet<String> = BTreeSet::new();
    for (_, daily) in per_sheet {
        all_dates.extend(daily.keys().cloned());
    }
    
    let mut columns = vec![DATE_COL.to_string()];
    columns.extend(sheet_names.iter().map(|n| (*n).clone()));
    columns.push("合计".to_string());
    
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut col_tots = vec![0usize; sheet_names.len()];
    let mut grand = 0usize;
    
    for date in &all_dates {
        let mut row = vec![date.clone()];
        let mut row_total = 0usize;
        for (i, (_, daily)) in per_sheet.iter().enumerate() {
            let val = daily.get(date.as_str()).map(|s| s.len()).unwrap_or(0);
            row.push(if val == 0 { String::new() } else { val.to_string() });
            row_total += val;
            col_tots[i] += val;
        }
        row.push(if row_total == 0 { String::new() } else { row_total.to_string() });
        grand += row_total;
        rows.push(row);
    }
    
    let mut totals = vec!["合计".to_string()];
    for &ct in &col_tots {
        totals.push(if ct == 0 { String::new() } else { ct.to_string() });
    }
    totals.push(if grand == 0 { String::new() } else { grand.to_string() });
    rows.push(totals.clone());
    
    TableData { columns, rows, totals, mode: mode.to_string() }
}

// ===================== Tauri 命令 =====================

#[tauri::command]
fn get_version() -> String { VERSION.to_string() }

#[tauri::command]
fn import_files(paths: Vec<String>, state: State<AppState>) -> Result<FileResult, String> {
    let mut fps = state.file_paths.lock().map_err(|e| e.to_string())?;
    let existing: HashSet<String> = fps.iter().cloned().collect();
    let new: Vec<String> = paths.into_iter().filter(|p| !existing.contains(p) && Path::new(p).exists()).collect();
    let cnt = new.len();
    fps.extend(new);
    let (files, table) = process_all_files(&fps);
    if let Ok(mut mr) = state.merged_results.lock() { *mr = Some(table.clone()); }
    app_log(&state, format!("导入文件：新增 {} 个，共 {} 个", cnt, fps.len()));
    Ok(FileResult { files, results: Some(table), status: format!("已导入 {} 个文件（新增 {} 个），统计完成", fps.len(), cnt), imported_count: cnt })
}

/// 递归收集目录下所有 .xlsx 文件（排除 ~$ 临时文件）
fn collect_xlsx(dir: &Path, files: &mut Vec<String>) {
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let p = entry.path();
            if p.is_dir() {
                collect_xlsx(&p, files);
            } else if p.extension().map(|x| x == "xlsx").unwrap_or(false)
                && !p.file_name().unwrap().to_string_lossy().starts_with("~$") {
                files.push(p.to_string_lossy().to_string());
            }
        }
    }
}

#[tauri::command]
fn import_folder(path: String, state: State<AppState>) -> Result<FileResult, String> {
    let dir = Path::new(&path);
    if !dir.is_dir() { return Err("无效的文件夹路径".to_string()); }
    let mut found = Vec::new();
    collect_xlsx(dir, &mut found);
    let mut fps = state.file_paths.lock().map_err(|e| e.to_string())?;
    let existing: HashSet<String> = fps.iter().cloned().collect();
    let new: Vec<String> = found.into_iter().filter(|p| !existing.contains(p)).collect();
    fps.extend(new.clone());
    let (files, table) = process_all_files(&fps);
    if let Ok(mut mr) = state.merged_results.lock() { *mr = Some(table.clone()); }
    app_log(&state, format!("导入文件夹：新增 {} 个文件，共 {} 个", new.len(), fps.len()));
    Ok(FileResult { files, results: Some(table), status: format!("导入完成：新增 {} 个文件（共 {} 个）", new.len(), fps.len()), imported_count: new.len() })
}

#[tauri::command]
fn process_files(state: State<AppState>) -> Result<FileResult, String> {
    let fps = state.file_paths.lock().map_err(|e| e.to_string())?;
    if fps.is_empty() { return Ok(FileResult { files: vec![], results: None, status: "没有文件需要处理".to_string(), imported_count: 0 }); }
    let paths = fps.clone(); drop(fps);
    let (files, table) = process_all_files(&paths);
    if let Ok(mut mr) = state.merged_results.lock() { *mr = Some(table.clone()); }
    app_log(&state, format!("统计完成：{} 个文件", paths.len()));
    Ok(FileResult { files, results: Some(table), status: format!("统计完成：{} 个文件", paths.len()), imported_count: 0 })
}

#[tauri::command]
fn get_results(file_index: i32, state: State<AppState>) -> TableData {
    let fps = state.file_paths.lock().map(|p| p.clone()).unwrap_or_default();
    if file_index < 0 || file_index as usize >= fps.len() {
        if let Some(data) = state.merged_results.lock().ok().and_then(|r| r.clone()) { return data; }
        return TableData { columns: vec![], rows: vec![], totals: vec![], mode: "全部文件 — 无数据".to_string() };
    }
    let path = &fps[file_index as usize];
    let filename = Path::new(path).file_name().unwrap().to_string_lossy().to_string();
    if let Some(per_sheet) = process_file_for_export(path) {
        return build_wide_table(&per_sheet, &filename);
    }
    TableData { columns: vec![], rows: vec![], totals: vec![], mode: format!("{} — 无数据", filename) }
}

/// Excel 列号转列字母（0-based: 0→A, 1→B, ... 25→Z, 26→AA）
fn col_letter(col: usize) -> String {
    let mut n = col;
    let mut s = String::new();
    loop {
        s.insert(0, char::from_u32(('A' as u32) + (n % 26) as u32).unwrap());
        if n < 26 { break; }
        n = n / 26 - 1;
    }
    s
}

/// 按 Sheet 分组的导出用处理函数，保留 Sheet 维度（与 Python 版一致）
/// 返回: sheet_name → date → batch_set
fn process_file_for_export(path: &str) -> Option<Vec<(String, BTreeMap<String, BTreeSet<String>>)>> {
    let mut wb = match calamine::open_workbook::<Xlsx<_>, _>(path) {
        Ok(wb) => wb,
        Err(_) => return None,
    };
    let mut result: Vec<(String, BTreeMap<String, BTreeSet<String>>)> = Vec::new();
    for name in wb.sheet_names().to_vec() {
        if name == "目录" { continue; }
        if let Ok(range) = wb.worksheet_range(&name) {
            if let Some(records) = process_sheet(&range) {
                let mut daily: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
                for (d, b) in records {
                    daily.entry(d).or_default().insert(b);
                }
                result.push((name, daily));
            }
        }
    }
    if result.is_empty() { None } else { Some(result) }
}

#[tauri::command]
fn batch_export(export_dir: String, state: State<AppState>) -> Result<String, String> {
    let fps = state.file_paths.lock().map_err(|e| e.to_string())?;
    if fps.is_empty() { return Err("没有文件可以导出".to_string()); }
    let paths = fps.clone(); drop(fps);
    if let Ok(mut ed) = state.export_dir.lock() { *ed = export_dir.clone(); }
    fs::create_dir_all(&export_dir).map_err(|e| format!("创建目录失败: {}", e))?;

    fn ch_w(s: &str) -> f64 { s.chars().map(|c| if c as u32 > 127 { 2.0 } else { 1.0 }).sum() }
    let fh = Format::new().set_font_name("仿宋").set_font_size(14).set_bold().set_border(FormatBorder::Thin).set_align(FormatAlign::Center).set_align(FormatAlign::VerticalCenter);
    let fd = Format::new().set_font_name("仿宋").set_font_size(14).set_border(FormatBorder::Thin).set_align(FormatAlign::Center).set_align(FormatAlign::VerticalCenter);
    let fb = Format::new().set_font_name("仿宋").set_font_size(14).set_bold().set_border(FormatBorder::Thin).set_align(FormatAlign::Center).set_align(FormatAlign::VerticalCenter);

    let mut success = 0u32;
    let mut failed: Vec<String> = Vec::new();
    for path in &paths {
        let base = Path::new(path).file_stem().unwrap().to_string_lossy().to_string();
        let out_path = Path::new(&export_dir).join(format!("{}_统计结果.xlsx", base));
        let per_sheet = match process_file_for_export(path) {
            Some(r) => r,
            None => { failed.push(Path::new(path).file_name().unwrap().to_string_lossy().to_string()); continue; }
        };
        let sheet_names: Vec<&String> = per_sheet.iter().map(|(name, _)| name).collect();
        let mut all_dates: BTreeSet<String> = BTreeSet::new();
        for (_, daily) in &per_sheet { all_dates.extend(daily.keys().cloned()); }
        let ncols = sheet_names.len() + 2;
        let total_col = (sheet_names.len() + 1) as u16;

        let mut wb = Workbook::new();

        // Sheet 1: 每日批号统计
        {
            let ws = wb.add_worksheet().set_name("每日批号统计").map_err(|e| e.to_string())?;
            ws.write_string_with_format(0, 0, DATE_COL, &fh).ok();
            let mut cw: Vec<f64> = vec![ch_w(DATE_COL); ncols];
            for (i, sn) in sheet_names.iter().enumerate() {
                ws.write_string_with_format(0, (i + 1) as u16, sn.as_str(), &fh).ok();
                cw[i + 1] = cw[i + 1].max(ch_w(sn));
            }
            ws.write_string_with_format(0, total_col, "合计", &fh).ok();
            cw[ncols - 1] = cw[ncols - 1].max(ch_w("合计"));

            let mut row = 1u32;
            for date in &all_dates {
                ws.write_string_with_format(row, 0, date, &fd).ok();
                cw[0] = cw[0].max(ch_w(date));
                let mut row_total = 0usize;
                for (i, (_, daily)) in per_sheet.iter().enumerate() {
                    let val = daily.get(date.as_str()).map(|s| s.len()).unwrap_or(0);
                    if val > 0 {
                        ws.write_number_with_format(row, (i + 1) as u16, val as f64, &fd).ok();
                        cw[i + 1] = cw[i + 1].max(ch_w(&val.to_string()));
                    } else {
                        ws.write_string_with_format(row, (i + 1) as u16, "", &fd).ok();
                    }
                    row_total += val;
                }
                if row_total > 0 {
                    ws.write_number_with_format(row, total_col, row_total as f64, &fd).ok();
                    cw[ncols - 1] = cw[ncols - 1].max(ch_w(&row_total.to_string()));
                } else {
                    ws.write_string_with_format(row, total_col, "", &fd).ok();
                }
                ws.set_row_height(row, 35.0).ok();
                row += 1;
            }

            // 合计行 SUM 公式（数据从第2行到第 data_end 行）
            let data_end = all_dates.len() as u32 + 1;
            ws.write_string_with_format(row, 0, "合计", &fb).ok();
            for i in 0..sheet_names.len() {
                let cl = col_letter(i + 1);
                ws.write_formula_with_format(row, (i + 1) as u16, format!("=SUM({cl}2:{cl}{data_end})").as_str(), &fb).ok();
            }
            let tcl = col_letter(sheet_names.len() + 1);
            ws.write_formula_with_format(row, total_col, format!("=SUM({tcl}2:{tcl}{data_end})").as_str(), &fb).ok();

            for i in 0..ncols { ws.set_column_width(i as u16, (cw[i] + 5.0).clamp(8.0, 30.0)).ok(); }
            ws.set_row_height(0, 35.0).ok();
            ws.set_row_height(row, 35.0).ok();
        }

        // Sheet 2: 每日汇总
        {
            let ws = wb.add_worksheet().set_name("每日汇总").map_err(|e| e.to_string())?;
            ws.write_string_with_format(0, 0, DATE_COL, &fh).ok();
            ws.write_string_with_format(0, 1, "合计", &fh).ok();
            let mut m0 = ch_w(DATE_COL); let mut m1 = ch_w("合计");

            let mut r2 = 1u32;
            for date in &all_dates {
                ws.write_string_with_format(r2, 0, date, &fd).ok();
                m0 = m0.max(ch_w(date));
                let dt: usize = per_sheet.iter().map(|(_, d)| d.get(date.as_str()).map(|s| s.len()).unwrap_or(0)).sum();
                if dt > 0 {
                    ws.write_number_with_format(r2, 1, dt as f64, &fd).ok();
                } else {
                    ws.write_string_with_format(r2, 1, "", &fd).ok();
                }
                ws.set_row_height(r2, 35.0).ok();
                r2 += 1;
            }
            let de2 = all_dates.len() as u32 + 1;
            ws.write_string_with_format(r2, 0, "合计", &fb).ok();
            ws.write_formula_with_format(r2, 1, format!("=SUM(B2:B{de2})").as_str(), &fb).ok();
            ws.set_column_width(0, (m0 + 5.0).clamp(8.0, 30.0)).ok();
            ws.set_column_width(1, (m1 + 5.0).clamp(8.0, 30.0)).ok();
            ws.set_row_height(0, 35.0).ok();
            ws.set_row_height(r2, 35.0).ok();
        }

        if wb.save(&out_path).is_err() {
            failed.push(Path::new(path).file_name().unwrap().to_string_lossy().to_string());
            continue;
        }
        success += 1;
    }
    app_log(&state, format!("批量导出：成功 {} / {} 个文件", success, paths.len()));
    if !failed.is_empty() {
        Ok(format!("成功导出 {} / {} 个文件，失败: {}", success, paths.len(), failed.join(", ")))
    } else {
        Ok(format!("成功导出 {} 个文件到目录: {}", success, export_dir))
    }
}

#[tauri::command]
fn get_export_dir(state: State<AppState>) -> String {
    state.export_dir.lock().map(|d| d.clone()).unwrap_or_default()
}

#[tauri::command]
fn open_export_folder(state: State<AppState>) {
    let dir = state.export_dir.lock().map(|d| d.clone()).unwrap_or_default();
    if !dir.is_empty() && Path::new(&dir).exists() { let _ = std::process::Command::new("explorer").arg(&dir).spawn(); }
}

#[tauri::command]
fn clear_all(state: State<AppState>) -> Result<(), String> {
    state.file_paths.lock().map_err(|e| e.to_string())?.clear();
    state.merged_results.lock().map_err(|e| e.to_string())?.take();
    app_log(&state, "清空所有数据".to_string());
    Ok(())
}

#[tauri::command]
fn get_log_content(state: State<AppState>) -> String {
    state.logs.lock().map(|l| l.join("\n")).unwrap_or_default()
}

#[tauri::command]
fn check_update() -> UpdateInfo {
    UpdateInfo { has_update: false, latest_version: VERSION.to_string(), current_version: VERSION.to_string(), release_notes: String::new(), download_url: String::new() }
}

// ===================== 应用入口 =====================

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let de = dirs::desktop_dir().map(|d| d.join("统计导出").to_string_lossy().to_string()).unwrap_or_else(|| "C:\\统计导出".to_string());
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .manage(AppState { file_paths: Mutex::new(Vec::new()), merged_results: Mutex::new(None), export_dir: Mutex::new(de), logs: Mutex::new(Vec::new()) })
        .invoke_handler(tauri::generate_handler![
            get_version, import_files, import_folder, process_files, get_results,
            batch_export, get_export_dir, open_export_folder,
            clear_all, get_log_content, check_update,
        ])
        .run(tauri::generate_context!())
        .expect("启动失败");
}
