<template>
  <main class="container">
    <div>
      1、
      <button @click="open_slx_file">选择数据文件xls、xlsx</button>
      表格文件路径：{{ xlsx_path }}
    </div>
    <div>2、选择工作表
      <select id="basicSelect" @change="select_table" v-model="table_name">
        <option v-for="name in table_name_list" :key="name" :value="name">
          {{ name }}
        </option>
      </select>
    </div>
    <div>3、择标题行：第<input type="number" v-model.number="title_row_number" @change="SetTitleS" class="inputnum" />行（请输入数字，0没有标题行）
    </div>

    <div>4、选择数据列：
      <label v-for="val in titleData" @change="TitleSelectFun" :key="val.name">
        <span :style="{ fontWeight: val.select ? 'bold' : 'normal', color: val.select ? 'blue' : 'black' }">
          {{ val.name }}
        </span>
        <input type="checkbox" v-model="val.select" class="checkboxtitle" />
      </label>
    </div>

    <div>5、选择数据行：第<input type="number" v-model.number="data_row_number_statr" class="inputnum" />行---第
      <input type="number" v-model.number="data_row_number_end" class="inputnum" />行
    </div>

    <div>6、
      <button @click="CreateDocxTemplateFun">生成docx简要模版</button>
    </div>
    <div>7、
      <button @click="open_template_file" disabled="false">选择其他docx模版(功能待完善)</button>
      模版文件路径：{{ template_path }}
    </div>
    <div>
      8、
      <button>生成一个docx文件（方便一次性打印）</button>
      <button>生成多个单独docx文件</button>
    </div>

  </main>
</template>

<script setup lang="ts">
import { ref, reactive } from "vue";
// import { invoke } from "@tauri-apps/api/core";
import { open } from '@tauri-apps/plugin-dialog';
import { dirname } from '@tauri-apps/api/path';
import { get_row_col, get_table_name_list, GetTitleS, CreateDocxTemplate } from "./useGetData.ts";
import { CreateTitleS, FileTimeStr, TitleSelectUtil } from "./utils.ts";

let xlsx_path = ref("");       // xls表格文件路径
let main_dir = ref("");        // xls表格文件路径所在文件夹
let out_dir = ref("");         // docx文件输出文件夹
let template_path = ref("");   // 使用的docx模版文件

let table_name_list = ref([]);            // 数据表名列表
let table_name = ref("");                // 当前选中的表
let title_row_number = ref(0);            // 标题所在的行
let data_row_number_statr = ref(0);       // 数据起始行
let data_row_number_end = ref(0);         // 数据结束行
let data_row_sum = ref(0);                  // 数据总行数
let data_col_sum = ref(0);                 // 数据总列数
let col_name = reactive([]);

let titleData = ref<Array<{ name: string, select: boolean }>>([]);   // xlsx标题行
let titleSelect = ref<Array<{}>>([]);   // 选中的标题列


// 选择表格文件
// 选择文件后，首先还原界面上的所有设置
// 选择xls文件后，读取文件绝对路径，和文件所在
// 读取xls，获取有效数据的行数、列数
async function open_slx_file() {
  const file = await open({
    multiple: false,               // 是否多选
    directory: false,              // 文件件是否可选
    filters: [{
      name: "表格文件",
      extensions: ["xls", "xlsx"]  // 只选择指定扩展名
    }]
  });
  if (file) {
    xlsx_path.value = file;                // 表格文件路径
    main_dir.value = await dirname(xlsx_path.value);  // 表格文件所在目录
    // 清空其他步骤数据
    titleSelect.value = [];    // 选中的标题列清空
    title_row_number.value = 0;       // 标题行归零
    data_row_number_statr.value = 1;  // 数据起始行归1
    table_name_list.value = await get_table_name_list(xlsx_path.value) as [];   // 获取表名列表
    if (table_name_list.value.length > 0) {
      table_name.value = table_name_list.value[0];                         // 2、中列表显示第一个表名
      [data_row_sum.value, data_col_sum.value] = await get_row_col(xlsx_path.value, table_name.value) as [number, number];  // 获取表中数据的行数、列数
      data_row_number_end.value = data_row_sum.value;     // 默认选择最大行数
      titleData.value = CreateTitleS(data_col_sum.value);        // 未选择标题行的时候，按照列名

    } else {
      table_name.value = "";
    }
  }
}

// 选择工作表
async function select_table() {
  title_row_number.value = 0;       // 标题行归零
  data_row_number_statr.value = 1;  // 数据起始行归1
  titleSelect.value = [];    // 选中的标题列清空
  console.log(xlsx_path.value, table_name.value);
  [data_row_sum.value, data_col_sum.value] = await get_row_col(xlsx_path.value, table_name.value) as [number, number];  // 获取表中数据的行数、列数
  console.log(data_row_sum.value, data_col_sum.value);
  data_row_number_end.value = data_row_sum.value;     // 默认选择最大行数
  titleData.value = CreateTitleS(data_col_sum.value);        // 未选择标题行的时候，按照列名

}

// 选择标题行
async function SetTitleS() {
  titleSelect.value = [];                     // 选中的标题列清空
  if (data_row_sum.value == 0) {                // 没有数据的时候
    title_row_number.value = 0;                 // 标题行=0
    data_row_number_statr.value = 0;            // 数据行起始=0
    data_row_number_end.value = 0;              // 数据行结束=0
    titleData.value = [];                       // 清空列名
    return
  }
  if (title_row_number.value <= 0 || data_row_sum.value == 1) {  // 标题行最小值为0，或者只有1行数据就没有标题行
    title_row_number.value = 0;                                  // 标题行=0
    data_row_number_statr.value = 1;                             // 数据航=1
    titleData.value = CreateTitleS(data_col_sum.value);          // 未选择标题行的时候，按照列名
  } else if (title_row_number.value >= data_row_sum.value) {     // 标题行最大值=总行数-1
    title_row_number.value = data_row_sum.value - 1;             // 标题行=总行数-1
    data_row_number_statr.value = data_row_sum.value;            // 数据行起始为最后一行
    data_row_number_end.value = data_row_sum.value;              // 数据行结束航为最后一行
    titleData.value = await GetTitleS(xlsx_path.value, table_name.value, title_row_number.value) as [];                           // 按照标题行赋值
  } else {                                                       // 范围内的时候
    data_row_number_statr.value = title_row_number.value + 1;    // 数据起始行=标题行+1
    data_row_number_end.value = data_row_sum.value;              // 数据结束航=总行数
    titleData.value = await GetTitleS(xlsx_path.value, table_name.value, title_row_number.value) as [];
  }
  // 当标题行不为0的时候，获取表格中标题所在行的数据
  // Object.assign(title_data, getTitleS(xlsx_path.value, table_name.value, title_row_number.value));
}

// 生成选中的标题字段
function TitleSelectFun() {
  titleSelect.value = TitleSelectUtil(titleData.value);
}

// 创建docx模版
async function CreateDocxTemplateFun() {
  // 模版文件路径“目录\模版20260206133706.docx”
  if ((titleSelect.value).length > 0) {
    template_path.value = `${main_dir.value}\\${FileTimeStr()}.docx`
    let result = await CreateDocxTemplate(template_path.value, titleSelect.value) as boolean;
    if (result) {
      // 创建成功，模版文件为。。。。
      console.log(result);
    }
  } else {
    alert("请先选择数据表列！");
  }

}

// 选择设置好的模版文件
async function open_template_file() {
  console.log("选择文件：---");
  // Open a dialog
  const file = await open({
    multiple: false,
    directory: false,
    defaultPath: main_dir.value,
    filters: [{
      name: "文档",
      extensions: ["docx"]  // 支持docx格式
    }]
  });
  if (file) {
    template_path.value = file;
  }
  // greetMsg.value = await invoke("greet", { name: name.value });
}
</script>

<style>
.container {
  margin: 0;
  padding-top: 2vh;
  display: flex;
  flex-direction: column;
  justify-content: center;
  text-align: left;
}

.container>div {
  margin-top: 2vh;
}

label {
  margin-left: 1vw;
}

.checkboxtitle {
  margin-left: 0vw;
}
.inputnum{
  width: 50px;
}
</style>