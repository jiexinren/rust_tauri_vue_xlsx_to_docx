use docx_rs::*;
use serde_json::{Map, Value};

/// 写模版文件
#[tauri::command]
pub fn create_docx_template(path: &str, title_select: Vec<Map<String, Value>>) -> String {
    let mut titles = String::new();
    for item in title_select {
        for (k, _) in item.iter() {
            titles = format!("{titles}{{{k}}}");
        }
    }

    let file = std::fs::File::create(path).unwrap(); // 创建文件
    let mut docx = Docx::new(); // 创建文档
    docx = docx.page_margin(
        docx_rs::PageMargin::new()
            .top(2097)
            .bottom(1984)
            .left(1587)
            .right(1474),
    );
    docx.add_paragraph(
        // 添加段落
        Paragraph::new() // 创建段落
            .line_spacing(LineSpacing::new().line(28)) // 设置行距
            .add_run(
                Run::new().add_text(&titles)
                .fonts(docx_rs::RunFonts::new().ascii("仿宋"))
                .size(24)
            ),
    ) // 将文本内容添加到段落
    .build() // 将Docx文档转换为XML文档
    .pack(file)
    .unwrap(); // 将XML文档压缩打包为docx文件
    titles
}
