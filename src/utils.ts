// 根据输入的n生成对应的标题数组
// 按照xlsx列号方式生成name
// [{"name":"A","selct":false},...]
export function CreateTitleS(n: number) {
    const result: Array<{ name: string, select: boolean }> = []
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    for (let i = 0; i < n; i++) {
        let title = '列'
        let num = i

        do {
            title = chars[num % 26] + title
            num = Math.floor(num / 26) - 1
        } while (num >= 0)

        result.push({ name: title, select: false })
    }
    console.log(result);
    return result
}

// 选中的标题列单独提取出来
export function TitleSelectUtil(titleData: Array<{ name: string, select: boolean }>) {
    let result: { [key: string]: number }[] = [];
    titleData.forEach((item, index) => {
        if (item.select) {
            result.push({ [item.name]: index })
        }
    });
    return result;
}

// 生成带时间的模版文件名
export function FileTimeStr(){
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    return `模版${year}${month}${day}${hours}${minutes}${seconds}`;
}
