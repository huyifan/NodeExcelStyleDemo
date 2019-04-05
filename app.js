const xlsx = require('node-xlsx').default;
const fs = require('fs')
const xlsx2 = require('xlsx')
const form = {
    name: '模拟数据表',
    data: [
        ['姓名', '性别', '年级', '单位', '政治面貌', '籍贯'],
        ['zhangsan', 'man1', '21', 'home1', 'people', 'china'],
        ['zhangsan1', 'man2', '22', 'home2', 'people', 'china'],
        ['zhangsan2', 'man3', '23', 'home3', 'people', 'china'],
        ['zhangsan3', 'man4', '24', 'home4', 'people', 'china'],
        ['zhangsan4', 'man5', '25', 'home5', 'people', 'china'],
        ['zhangsan5', 'man6', '26', 'home6', 'people', 'china'],
        ['zhangsan6', 'man7', '27', 'home7', 'people', 'china'],
    ]
}


const form1 = {
    name: '模拟数据表',
    data: [
        ['姓名', '性别', '年级', '单位', '政治面貌', '籍贯'],
        [
            {
                v: 'zhangsan',
                s: {
                    font: {
                        size: 19,
                        bold: true,
                        color: {rgb: 'f33004'}
                    }

                }
            }
            ,
            'man1', '21', 'home1', 'people', 'china'],
        [{
            v: 'zhangsan1',
            s: {
                font: {
                    size: 19,
                    bold: true,
                    color: {rgb: '573dff'}
                },
                fill: {
                    fgColor: {
                        rgb: 'ffff00'
                    }
                },


            }
        }
            , 'man2', '22', 'home2', 'people', 'china'],
        ['zhangsan2', 'man3', '23', 'home3', 'people', 'china'],
        ['zhangsan3', 'man4', '24', 'home4', 'people', 'china'],
        ['zhangsan4', 'man5', '25', 'home5', 'people', 'china'],
        ['zhangsan5', 'man6', '26', 'home6', 'people', 'china'],
        ['zhangsan6', 'man7', '27', 'home7', 'people', 'china'],
    ]
}

const mockData = []
const form2 = {
    name: '认真的表格',
    data: mockData
}

form.data.map((v, i) => {
    if (i === 0) {
        const firstLine = []

        v.map((firstItem, i) => {
            firstLine.push({
                v: firstItem,
                s: {
                    alignment: {
                        vertical: 'center',
                        horizontal: 'center'
                    },
                    font: {
                        size: 19,
                        bold: true,
                        color: {rgb: 'ffffff'}
                    },
                    fill: {
                        fgColor: {
                            rgb: 'a4a3a5'
                        }
                    }
                }
            })
        })

        mockData.push(firstLine)

    } else {
        const line = []
        v.map((item, i) => {
            line.push({
                v: item,
                s: {
                    alignment: {
                        vertical: 'center',
                        horizontal: 'center'
                    },
                    font: {
                        size: 19,
                        color: {rgb: 'ff280c'}
                    }
                }
            })
        })
        mockData.push(line)

    }


})


const options = {
    '!cols': [ //设置宽度
        {wpx: 100},//1-姓名
        {wpx: 140},//2-性别
        {wpx: 180},//3-年级
        {wpx: 220}, //4-单位
        {wpx: 260}, //5-政治面貌
        {wpx: 300}, //6-籍贯

    ],
    //高度设置无效
    '!rows': [//设置高度
        {hpx: 40,}, //1
        {hpx: 60},//2
        {hpx: 80},//3
        {hpx: 100},//4
        {hpx: 120},//5
        {hpx: 120},//6
        {hpx: 120},//7
        {hpx: 120},//8
        {hpx: 120},//9
    ],
    '!margins': {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
}

// const range = {s: {c: 0, r: 0}, e: {c: 0, r: 2}}; // A1:A4
// options['!merges'] = [range]


const xlsxData = xlsx.build([form2], options)


console.log("准备写入文件");
fs.writeFile('input.xlsx', xlsxData, function (err) {
    if (err) {
        return console.error(err);
    }
    console.log("数据写入成功！");
    console.log("--------我是分割线-------------")

});