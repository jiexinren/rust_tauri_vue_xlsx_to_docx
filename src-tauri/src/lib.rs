mod func;
mod my_docx;
mod my_xlsx;

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            my_xlsx::get_table_name_list,  // 获取数据表中的表名列表
            my_xlsx::get_row_col,          // 获取有效数据总行数和列数
            my_xlsx::get_title,            // 获取标题行标题名
            my_docx::create_docx_template  // 创建doc模版文件
        ])
        .run(tauri::generate_context!())
        .expect("tauri启动错误！");
}
