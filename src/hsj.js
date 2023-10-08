import * as GC from "@grapecity/spread-sheets";
import "@grapecity/spread-sheets-io"
import "@grapecity/spread-sheets-shapes"
import "@grapecity/spread-sheets-resources-zh"
import "@grapecity/spread-sheets-designer-resources-cn"
import "@grapecity/spread-sheets-designer"

GC.Spread.Common.CultureManager.culture("zh-cn");
var oldFun = GC.Spread.Sheets.getTypeFromString;

GC.Spread.Sheets.getTypeFromString = function (typeString) {
    switch (typeString) {
        case "cellType1":
            return cellType1;
        case "cellType2":
            return cellType2;
        default:
            return oldFun.apply(this, arguments);
    }
};




let designerConfig = JSON.parse(JSON.stringify(GC.Spread.Sheets.Designer.DefaultConfig))
let customerRibbon = {
    "id": "operate",
    "text": "自定义菜单",
    "buttonGroups": [
    ]
};

let ribbonFileConfig = [
    {
        "label": "操作类型",
        "thumbnailClass": "ribbon-thumbnail-spreadsettings",
        "commandGroup": {
            "children": [
                {
                    "commands": ["setCellType1", "setCellType2"]
                }
            ]
        }
    },
    {
        "label": "设定区域",
        "thumbnailClass": "ribbon-thumbnail-spreadsettings",
        "commandGroup": {
            "children": [
                {
                    "commands": ["setMoveArea1", "setMoveArea2"]
                }
            ]
        }
    },
    {
        "label": "设置数据源",
        "thumbnailClass": "ribbon-thumbnail-spreadsettings",
        "commandGroup": {
            "children": [
                {
                    "commands": ["initDataSource"]
                }
            ]
        }
    }
]

// 定义命令
let ribbonFileCommands = {
    "setCellType1": {
        iconClass: "ribbon-button-welcome", // 按钮样式，可以指定按钮图片
        text: "类型1", // 显示文本
        commandName: "setCellType1", // 命令名称，需要全局唯一
        execute: async function (context) { // 命令对应的操作
            setCellType_1(context)
        }
    },
    "setCellType2": {
        iconClass: "ribbon-button-welcome",
        text: "类型2",
        commandName: "setCellType2",
        execute: async function (context) {
            setCellType_2(context)
        }
    },
    "setMoveArea1": {
        iconClass: "ribbon-button-welcome",
        text: "需要移动的区域1",
        commandName: "setMoveArea1",
        execute: async function (context) {
            setMoveArea_1(context)
        }
    },
    "setMoveArea2": {
        iconClass: "ribbon-button-welcome",
        text: "需要移动的区域2",
        commandName: "setMoveArea2",
        execute: async function (context) {
            setMoveArea_2(context)
        }
    },
    "initDataSource": {
        iconClass: "ribbon-button-welcome",
        text: "设置数据源",
        commandName: "initDataSource",
        execute: async function (context) {
            initDataSource(context)
        }
    },
}

// 注册命令到config的commandMap属性上
designerConfig.commandMap = {};
Object.assign(designerConfig.commandMap, ribbonFileCommands);

// 在designer的config中加入自定义的命令和按钮
customerRibbon.buttonGroups = customerRibbon.buttonGroups.concat(ribbonFileConfig);
designerConfig.ribbon.push(customerRibbon);

designer = new GC.Spread.Sheets.Designer.Designer("designer-container", designerConfig)
spread = designer.getWorkbook()





let oldTextPaint = GC.Spread.Sheets.CellTypes.Text.prototype.paint

function cellType1() {
    this.typeName = "cellType1"
}
cellType1.prototype = new GC.Spread.Sheets.CellTypes.Text()
cellType1.prototype.paint = function (canvas, value) {
    value = value || {
        notify: {}
    }
    this.value = value
    if (value && typeof value == "object") {
        let text =
`   我们于  ${value.notify.year || ""}  年 ${value.notify.month || ""}  月 ${value.notify.day || ""}  日收到了《义务教育入学通知书》，并认真阅读了家长须知。
    现按要求，请送子女（或被监护人）  ${value.name || ""}  到  ${value.school || ""}  报道。`
        this.text = text
        let arg = arguments
        arg[1] = text
        oldTextPaint.apply(this, arg)
    } else {
        oldTextPaint.apply(this, arguments)
    }
}

function cellType2() {
    this.typeName = "cellType2"
}
cellType2.prototype = new GC.Spread.Sheets.CellTypes.Text()
cellType2.prototype.paint = function (canvas, value) {
    value = value || {
        notify: {}
    }
    this.value = value
    if (value && typeof value == "object") {
        let text =
`   家长（监护人）  ${value.parent || ""}  ：
    您的子女（被监护人）  ${value.name || ""}  今年九月一日前达到本地规定的义务教育入学年龄。根据《中华人民共和国义务教育法》，您应送其入学接受义务教育。请于  ${value.notify.year || ""}  年 ${value.notify.month || ""}  月 ${value.notify.day || ""}  日，将其送到  ${value.school || ""}  学校就读，并保证完成九年义务教育。`
        this.text = text
        let arg = arguments
        arg[1] = text
        oldTextPaint.apply(this, arg)
    } else {
        oldTextPaint.apply(this, arguments)
    }
}


function setCellType_1(designer) {
    let spread = designer.getWorkbook()
    let sheet = spread.getActiveSheet()
    let { row, col } = sheet.getSelections()[0]
    let tag = sheet.tag() || {}
    tag.type1 = {
        position: {
            row: row,
            col: col,
        },
        moveArea: tag.type1?.moveArea || {}
    }
    sheet.tag(tag)
    sheet.setBindingPath(row, col, "data")
    sheet.setCellType(row, col, new cellType1())
    alert("设置成功！")
}

function setCellType_2(designer) {
    let spread = designer.getWorkbook()
    let sheet = spread.getActiveSheet()
    let { row, col } = sheet.getSelections()[0]
    let tag = sheet.tag() || {}
    tag.type2 = {
        position: {
            row: row,
            col: col,
        },
        moveArea: tag.type2?.moveArea || {}
    }
    sheet.tag(tag)
    sheet.setBindingPath(row, col, "data")
    sheet.setCellType(row, col, new cellType2())
    alert("设置成功！")
}

function setMoveArea_1(designer) {
    let spread = designer.getWorkbook()
    let sheet = spread.getActiveSheet()
    let { row, col, rowCount, colCount } = sheet.getSelections()[0]
    let tag = sheet.tag() || {}
    tag.type1 = {
        position: tag.type1?.position || {},
        moveArea: {
            row: row,
            col: col,
            rowCount: rowCount,
            colCount: colCount
        }
    }
    sheet.tag(tag)
    alert("设置成功！")
}
function setMoveArea_2(designer) {
    let spread = designer.getWorkbook()
    let sheet = spread.getActiveSheet()
    let { row, col, rowCount, colCount } = sheet.getSelections()[0]
    let tag = sheet.tag() || {}
    tag.type2 = {
        position: tag.type2?.position || {},
        moveArea: {
            row: row,
            col: col,
            rowCount: rowCount,
            colCount: colCount
        }
    }
    sheet.tag(tag)
    alert("设置成功！")
}

function initDataSource(designer) {
    let mock_data = {
        data: {
            street: "西安市xxxx街道",
            community: "xxxx社区",
            name: "小明",
            gender: "男",
            id: "xxxxxxxxxxxxxxxxxxxxxx",
            birthday: "2018.8.8",
            parent: "大明",
            school: "西安市中心实验小学",
            notify: {
                year: 2023,
                month: 7,
                day: 15
            }
        }
    }
    console.log(mock_data)
    let spread = designer.getWorkbook()
    let sheet = spread.getActiveSheet()
    let source = new GC.Spread.Sheets.Bindings.CellBindingSource(mock_data)
    sheet.setDataSource(source)
    sheet.scroll(-10000, -10000)
    adjustAreaHeight("type1", sheet)
    adjustAreaHeight("type2", sheet)

}

// 调整合并区域的范围
function adjustAreaHeight(typeName, sheet) {
    // 计算这一列的宽度
    let { row, col } = sheet.tag()[typeName].position
    let moveArea = sheet.tag()[typeName].moveArea
    let areaSpan = sheet.getSpan(row, col)
    let columnWidth = 0
    // 因为这里是一个合并区域，需要计算所有合并的列然后求和
    for (let i = areaSpan.col; i < areaSpan.col + areaSpan.colCount; i++) {
        columnWidth = columnWidth + sheet.getColumnWidth(i)
    }
    // 计算真实的文本长度
    let textWidth = measureWidth(sheet, row, col)
    // 真实需要的行数
    let needLine = Math.floor(textWidth / columnWidth) + Math.ceil((textWidth % columnWidth) / columnWidth)
    // diff就是目前需要的行数与实际行数之间的差别
    let diff = needLine - areaSpan.rowCount
    sheet.moveTo(moveArea.row, moveArea.col, moveArea.row + diff, moveArea.col, moveArea.rowCount, moveArea.colCount, GC.Spread.Sheets.CopyToOptions.all)
    sheet.removeSpan(areaSpan.row, areaSpan.col)
    sheet.addSpan(areaSpan.row, areaSpan.col, areaSpan.rowCount + diff, areaSpan.colCount)
}

// 计算真实的文本长度
function measureWidth(sheet, row, col) {
    let dom = document.createElement("canvas")
    let canvas = dom.getContext("2d")
    canvas.font = sheet.getStyle(row, col)?.font || "11pt Calibri"
    // 计算文本的宽度
    let text_width = canvas.measureText(sheet.getCell(row, col).cellType().text).width
    return text_width
}


let xhr = new XMLHttpRequest()
xhr.open("get", "./t.sjs")
xhr.responseType = "blob"
xhr.addEventListener("loadend", function () {
    if (this.readyState == 4 && this.status == 200) {
        let file = new File([this.response], 'test.sjs')
        spread.open(file)
    }
})
xhr.send()
