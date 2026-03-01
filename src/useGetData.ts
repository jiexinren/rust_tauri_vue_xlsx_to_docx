import {invoke} from "@tauri-apps/api/core";


// 根据数据表的绝对路径、工作表名，获取数据表中对应工作表中，数据的行数和列数
export async function get_row_col(path: string, tableName: string) {
    try {
        const result = await invoke("get_row_col", {
            path,
            tableName,                     // 前端变量名tableName对应rust后端变量名table_name
        });
        return result;
    } catch (error) {
        alert("获取行列数，加载错误！");
        return [];
    }
}

// 根据数据表绝对路径，获取表中所有的工作表名
export async function get_table_name_list(path: string) {
    try {
        const result = await invoke("get_table_name_list", {path});
        return result;
    } catch (error) {
        alert("获取表名列表，加载错误!");
        return [];
    }
}

// 根据数据表的绝对路径、工作表名，获取数据表中对应工作表中，标题行的标题名
export async function GetTitleS(path: string, tableName: string, titleRow: number) {
    try {
        const result = await invoke("get_title", {
            path,                          // 数据表绝对路径
            tableName,                     // 工作表明
            titleRow,                      // 标题所在行
        });
        console.log(result);
        return result;
    } catch (error) {
        alert("获取标题行，加载错误！");
        return [];
    }
}

// 根据模版文件路径，标题行字段数组，创建简单模版，成功返回true
export async function CreateDocxTemplate(path: string, titleSelect: {}[]){
    try {
        const result = await invoke("create_docx_template", {
            path,                            // 模版文件绝对路径
            titleSelect,                     // 标题字段数组
        });
        return result;
    } catch (error) {
        alert("创建模版文件，加载错误！");
        return [];
    }
}