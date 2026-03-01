use calamine::{open_workbook_auto, HeaderRow, Reader};
use serde::{Deserialize, Serialize};

/// 获取数据表名列表
#[tauri::command]
pub fn get_table_name_list(path: &str) -> Vec<String> {
    // 自动识别格式
    let workbook = open_workbook_auto(path).expect(&format!("读取:{},错误!", path));
    workbook.sheet_names()
}

/// 获取表格中数据的行数和列数
#[tauri::command]
pub fn get_row_col(path: &str, table_name: &str) -> Vec<i32> {
    // 打开文件
    let mut workbook = open_workbook_auto(path).expect(&format!("读取:{},错误!", path));
    // 打开工作表
    let sheet = workbook
        .with_header_row(HeaderRow::Row(0))
        .worksheet_range(table_name)
        .expect("读取工作表失败！");
    // 获取行数
    let (row, col) = sheet.get_size();
    vec![row as i32, col as i32]
}
#[derive(Deserialize, Serialize, Debug)]
pub struct Title {
    name: String,
    select: bool,
}

/// 获取标题行标题名
#[tauri::command]
pub fn get_title(path: &str, table_name: &str, title_row: i32) -> Vec<Title> {
    // 打开文件
    let mut workbook = open_workbook_auto(path).expect(&format!("读取:{},错误!", path));
    // 打开工作表
    let sheet = workbook
        .with_header_row(HeaderRow::Row(0))
        .worksheet_range(table_name)
        .expect("读取工作表失败！");
    // 获取标题行标题名
    let mut title = Vec::new();
    let (_, col_count) = sheet.get_size();

    for col in 0..col_count {
        if let Some(cell) = sheet.get(((title_row-1) as usize, col)) {
            title.push(Title {
                name: cell.to_string(),
                select: false,
            });
        }
    }
    title
}
