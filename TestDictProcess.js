var InputExcelArray = [
    ["DPP_RM_01", "DPP Ream", "OWN", "Domestic", "", "", "", "", "x", "x", "", "", "x", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RM_02", "DPP Ream", "OWN", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RL_01", "DPP Roll", "OWN", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "x", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RL_02", "DPP Roll", "OWN", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "x", "", "", "", "thanidao", "svd@scg.com", ""],
    ["PBL_RL_01", "Gypsum Roll", "OWN", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "x", "", "", "", "thanidao", "svd@scg.com", ""],
    ["PBL_RL_02", "Gypsum Roll", "OWN", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "x", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RM_03", "DPP Ream", "OS", "Domestic", "", "", "", "", "x", "x", "", "", "x", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RM_04", "DPP Ream", "OS", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RL_03", "DPP Roll", "OS", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "", "", "thanidao", "svd@scg.com", ""],
    ["DPP_RL_04", "DPP Roll", "OS", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "", "", "thanidao", "svd@scg.com", ""],
    ["PBL_RL_03", "Gypsum Roll", "OS", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "", "", "thanidao", "svd@scg.com", ""],
    ["PBL_RL_04", "Gypsum Roll", "OS", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "", "", "thanidao", "svd@scg.com", ""],
    ["KRT_RL_01", "Kraft Roll", "OS (Out-Out)", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "x", "", "x", "x", "x", "", "", "", "x", "", "", "thanidao", "svd@scg.com", ""],
    ["KRT_RM_01", "Kraft Ream", "OS", "Domestic", "", "", "", "", "", "x", "", "", "x", "", "", "", "x", "", "x", "", "", "x", "", "x", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["BAG_01", "Bag", "OS", "Domestic", "", "", "", "", "", "x", "", "", "x", "", "", "", "x", "", "x", "", "", "x", "", "x", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["SLT_01", "Slit", "OS", "Domestic", "", "", "", "", "", "x", "", "", "x", "", "", "", "x", "", "x", "", "", "x", "", "x", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["SLT_02", "Slit", "OS", "Export", "", "", "", "", "", "x", "", "", "", "", "", "", "x", "", "x", "", "", "x", "", "x", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["IBB_01", "IBB", "OWN", "Domestic", "", "", "", "", "", "x", "", "", "x", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["IBB_02", "IBB", "OWN", "Export", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "x", "", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["SHB_01", "Sheet Board", "OS", "Domestic", "", "", "", "", "", "x", "x", "x", "", "x", "", "x", "x", "", "x", "", "", "x", "", "x", "", "", "o", "", "jitsupan", "svd@scg.com", ""],
    ["KRT_RL_02", "Kraft Roll", "OS (In-Out)", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "sangdeac", "svd@scg.com", ""],
    ["KRT_RL_03", "Kraft Roll", "OS (Out-In)", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "x", "", "sangdeac", "svd@scg.com", ""],
    ["KRT_RL_04", "Kraft Roll", "OS (In-Out)", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "", "x", "", "sangdeac", "svd@scg.com", ""],
    ["KRT_RL_05", "Kraft Roll", "OWN", "Domestic", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "", "", "sangdeac", "svd@scg.com", ""],
    ["KRT_RL_06", "Kraft Roll", "OWN", "Export", "", "", "", "", "x", "x", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "", "sangdeac", "svd@scg.com", ""],
    ["FC_RL_01", "FC Roll", "OWN", "", "A01", "X", "X", "<> K000", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_RM_01", "FC Ream", "OWN", "", "A02", "X", "X", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_PP_01", "FC Pulp", "OWN", "", "", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_RL_02", "FC Roll", "OS", "", "A01", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_RM_02", "FC Ream", "OS", "", "A02", "", "", "", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "", "x", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_RF_01", "FC Ream Fest", "OS", "", "A03", "", "", "", "", "x", "", "", "x", "", "x", "", "", "", "", "", "", "", "x", "", "", "x", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_RF_02", "FC Ream Fest", "OWN", "", "A03", "X", "X", "", "", "x", "", "", "x", "", "x", "", "", "", "", "", "", "", "x", "", "x", "", "", "x", "supaporw", "svd@scg.com", ""],
    ["FC_SNP_01", "FC SNP", "OWN", "", "A01", "X", "X", "K000", "", "x", "", "", "", "", "", "", "", "", "", "", "", "", "x", "", "x", "", "", "x", "supaporw", "svd@scg.com", ""]
]

var fail = ["STOP","STOP","STOP","STOP","GO","GO","GO","GO","GO","GO","GO","GO","GO","GO","GO","GO","STOP","STOP","STOP","GO"];

var New2DArray = [];
New2DArray.push(pushFail("Global", "Fail", fail));

for (let Line of InputExcelArray) {
    let key = "";
    let value = "";
    let array = [];

    New2DArray.push(addKeyArrayValue(Line[0], "Process", Line));
    New2DArray.push(addKeyValue(Line[0], "Product", Line[1]));
    New2DArray.push(addKeyValue(Line[0], "OWN_OS", Line[2]));
    New2DArray.push(addKeyValue(Line[0], "Channel", Line[3]));
    New2DArray.push(addKeyValue(Line[0], "MatGroup1", Line[4]));
    New2DArray.push(addKeyValue(Line[0], "MRP", Line[5]));
    New2DArray.push(addKeyValue(Line[0], "WorkScheduling", Line[6]));
    New2DArray.push(addKeyValue(Line[0], "REMProfile", Line[7]));
    New2DArray.push(addKeyValue(Line[0], "RequestBy", Line[28]));
    New2DArray.push(addKeyValue(Line[0], "ToEmail", Line[29]));
    New2DArray.push(addKeyValue(Line[0], "CCEmail", Line[30]));
}

function addKeyValue(ProcessID, Key, Value) {
    let array = [];
    array.push(ProcessID + "_" + Key);
    array.push(Value);
    return array;
}

function addKeyArrayValue(ProcessID, Key, Line) {
    let array = [];
    array.push(ProcessID + "_" + Key);
    array.push(pushValuetoArray(Line));
    return array;
}

function pushValuetoArray(Line){
    let valueArray = [];
    for(let i = 8; i <= 27; i++){
        if (i == 26) {
            valueArray.push(Line[i].toLowerCase());
        } else {
            if (Line[i].toLowerCase() == 'x'){
                valueArray.push(true);
            } else{
                valueArray.push(false);
            }
        }
    }
    return valueArray;
}

function pushFail(ProcessID, Key, Line){
    let array = [];
    let valueArray = [];
    for(let element of Line){
        if (element.toLowerCase() == 'go'){
            valueArray.push(true);
        } else{
            valueArray.push(false);
        }
    }
    array.push(ProcessID + "_" + Key);
    array.push(valueArray);

    return array;
}

// Known as Automation "Create Dict from 2D Array"
var Dict = Object.fromEntries(New2DArray);

console.log(Dict["Global_Fail"]);
console.log(Dict["DPP_RM_01_Process"]);
console.log(Dict);
