#!/usr/bin/env node 
const ps = new PerfectScrollbar('#cells', {
    wheelSpeed: 12,
    wheelPropagation: true
});


function colIdtoName(n){
    let str = " ";

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = 'Z' + str; //'Z' is represented ny 0
            n = Math.floor((n / 26)) - 1;
        }
        else {
            str = String.fromCharCode(rem - 1 + 65) + str; //ascii to string
            n = Math.floor((n / 26));
        }
    }

    return str;

}

for (let i = 1; i <= 100; i++) {  
    let str=colIdtoName(i);
    $(`#coloumns`).append(`<div class="coloumn-name">${str}</div>`);
    $(`#rows`).append(`<div class="row-name">${i}</div>`);
}

// for scrolling rows and coloumns with cells 
//for this overflow hidden property is compulsory
$("#cells").scroll(function () {
    $("#coloumns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
})

function getIDs(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

function getAllFourCells(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId - 1}-col-${colId + 1}`);
    return [topCell, bottomCell, leftCell, rightCell];
}

let cellData = { "Sheet1": {} }; //cellData={ Sheet1{row{col{data}}}}
let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastlyAddedSheet = 1;//helpful in renaming 
let mouseMoved = false;
let startCellStored = false;
let startCell;
let endCell;
let saved=true;
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": 14,
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "color": "#444",
    "bgcolor": "#fff",
    "formula": "",
    "upStream":[],
    "downStream":[]
}

function loadNewSheet() {
    $("#cells").text("");
    for (let i = 1; i <= 100; i++) {
        let row = $(`<div class="cell-row"></div>`); //row bnai
        for (let j = 1; j <= 100; j++)
            row.append(`<div id= "row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);   //ik cell bnaya
        $(`#cells`).append(row);
    }
    addEventsToCells();
    addTabClickEvents();
}

loadNewSheet(); //for sheet 1

function getCellsInRange(start, end) {

    for (let i = (start.rowId < end.rowId ? start.rowId : end.rowId); i <= (start.rowId < end.rowId ? end.rowId : start.rowId); i++) {
        for (let j = (start.colId < end.colId ? start.colId : end.colId); j <= (start.colId < end.colId ? end.colId : start.colId); j++) {
            let [topCell, bottomCell, leftCell, rightCell] = getAllFourCells(i, j);
            selectCell($(`#row-${i}-col-${j}`)[0], {}, topCell, bottomCell, leftCell, rightCell, true);
        }
    }
}

function addEventsToCells() {
    //on double clicking a cell it is focused
    $(".input-cell").dblclick(function () {
        $(this).attr("contenteditable", "true");
        $(this).focus();
    })

    $(".input-cell").blur(function () {
        $(this).attr("contenteditable", "false");
        //let [rowId,colId]=getIDs(this);
        //cellData[selectedSheet][rowId-1][colId-1].text=$(this).text();
        updateCellData("text", $(this).text());
    })

    //on clicking a cell it is selected or deselected
    $(".input-cell").click(function (e) {
        let [rowId, colId] = getIDs(this);
        
        let [topCell, bottomCell, leftCell, rightCell] = getAllFourCells(rowId, colId);
        if ($(this).hasClass("selected") && e.ctrlKey) { //for multiple deselection

            deSelectCell(this, e, topCell, bottomCell, leftCell, rightCell)
        }
        else {
            selectCell(this, e, topCell, bottomCell, leftCell, rightCell) //it itself handles both cases of single and multiple selection
            //let formula= cellData[selectedSheet][rowId-1][colId-1]["formula"];
            //if(formula!="")
             //
        }
    })

    $(".input-cell").mousemove(function (event) {
        if (event.buttons == 1) { //if mouse is left clicked 
            event.preventDefault(); //prevent by default events
            $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
            mouseMoved = true;
            if (!startCellStored) //if abhi hold krna strt kia
            {
                startCellStored = true;
                let [rowId, colId] = getIDs(event.target);
                startCell = { rowId: rowId, colId: colId };
            }
            else {  //if already selected h kuch cells mtlb hold h
                let [rowId, colId] = getIDs(event.target);
                endCell = { rowId: rowId, colId: colId };
                getCellsInRange(startCell, endCell);
            }
        }
        else if (event.buttons == 0 && mouseMoved) { //for selecting something else afterwards
            startCellStored = false;
            mouseMoved = false;
        }
    })
}


function deSelectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {

    if (e.ctrlKey && $(ele).attr("contenteditable") == "false") {
        if ($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }
        if ($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }
        if ($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }
        if ($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }
        $(ele).removeClass("selected right-selected left-selected top-selected bottom-selected");
    }
}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mouseSelection) {

    if (e.ctrlKey || mouseSelection) {  //for multiple selections

        let topSelected;
        if (topCell) topSelected = topCell.hasClass("selected"); //check if its top cell is also selected

        let bottomSelected;
        if (bottomCell) bottomSelected = bottomCell.hasClass("selected");

        let leftSelected;
        if (leftCell) leftSelected = leftCell.hasClass("selected");

        let rightSelected;
        if (rightCell) rightSelected = rightCell.hasClass("selected");

        //if true than add bottom selected class to topcell which will make its bottom border gray 
        //and add top slected class to itself which will remove its top border
        if (topSelected) {
            topCell.addClass("bottom-selected");
            $(ele).addClass("top-selected");
        }
        if (bottomSelected) {
            bottomCell.addClass("top-selected");
            $(ele).addClass("bottom-selected");
        }
        if (rightSelected) {
            rightCell.addClass("left-selected");
            $(ele).addClass("right-selected");
        }
        if (leftSelected) {
            leftCell.addClass("right-selected");
            $(ele).addClass("left-selected");
        }
    }
    //if single cell selected first remove all classes in case of earlier selecteion
    else {
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
    }
    let [rowId,colId]=getIDs(ele);
    $(ele).addClass("selected");//add seleccted
    changeHeader([rowId,colId]);
}

//works whenever i select a class irrespective of my selection of style
function changeHeader([rowId, colId]) {
    let data;
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1])//agr exust krta h mtlb proprty chnge hui h
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    else data = defaultProperties;
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family", data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#color-fill-icon").css("border-bottom", `3px solid ${data.bgcolor}`);
    $("#text-fill-icon").css("border-bottom", `3px solid ${data.color}`);
    $("#function-input").text(data["formula"]);
}

function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    }
    else
        $(`#${property}`).removeClass("selected");
}

function setFontStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        //     $(".input-cell.selected").each(function(index,data){
        //       //let [rowId,colId]=getIDs(data);
        //       //cellData[selectedSheet][rowId-1][colId-1][property]=false;
        //     
        //   })
        updateCellData(property, false);
    } else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        //   $(".input-cell.selected").each(function(index,data){
        //       //let [rowId,colId]=getIDs(data);
        //       //cellData[selectedSheet][rowId-1][colId-1][property]=true;

        //   })
        updateCellData(property, true);
    }
}

$("#bold").click(function (e) {
    setFontStyle(this, "bold", "font-weight", "bold");
})

$("#italic").click(function (e) {
    setFontStyle(this, "italic", "font-style", "italic")
})

$("#underlined").click(function (e) {
    setFontStyle(this, "underlined", "text-decoration", "underline");
})

$(".menu-selector").change(function (e) {
    let value = $(this).val();
    key = $(this).attr("id");
    if (key == "font-family") {
        $("#font-family").css(key, value);
    }
    if (!isNaN(value)) {
        value = parseInt(value); //for font size
    }
    $(".input-cell.selected").css(key, value);
    //    $(".input-cell.selected").each(function(index,data){
    //        //let [rowId,colId]=getIDs(data);
    //        //cellData[selectedSheet][rowId-1][colId-1][key]=value;

    //    });
    updateCellData(key, value);

})

$(".alignment").click(function (e) {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");

    $(".input-cell.selected").css("text-align", alignment);
    // $(".input-cell.selected").each(function(index,data){
    //     //let [rowId,colId]=getIDs(data);
    //     cellData[selectedSheet][rowId-1][colId].alignment = alignment;
    // })
    updateCellData("alignment", alignment);
})

function updateCellData(property, value) {
    let prevCellData= JSON.stringify(cellData);
    if (value != defaultProperties[property]) { //alignment="left"
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getIDs(data);
            if (cellData[selectedSheet][rowId - 1] == undefined) {
                cellData[selectedSheet][rowId - 1] = {};
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream":[], "downStream":[]};//copy bnake dalre hai else defaul m hi change ho jaega
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
            } else {
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) {
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream":[], "downStream":[] };//copy bnake dalre hai else defaul m hi change ho jaega
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                } else {
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
            }
        });
    } else { //if firse default property p hi click kia h
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = getIDs(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) { //but uska data exist krta h mtlb changes h
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                if (JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][rowId - 1][colId - 1];
                    if (Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) //agr cell ka data delte krne k baad row{} empty ho jae to delete it
                        delete cellData[selectedSheet][rowId - 1];
                }
            }
        });  
    }
    if(saved && JSON.stringify(cellData) != prevCellData) 
        saved = false;

}

$(".color-pick").colorPick({
    'initialColor': "#TYPECOLOR",
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],

    'onColorSelected': function () {

        if (this.color != "#TYPECOLOR") {
            if (this.element.attr("id") == "fill-color") {
                $("#color-fill-icon").css("border-bottom", `3px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                // $(".input-cell.selected").each((index,data) => {
                //     let [rowId,colId] = getIDs(data);
                //     cellData[selectedSheet][rowId-1][colId-1].bgcolor = this.color;
                // });
                updateCellData("bgcolor", this.color);
            } else {
                $("#text-fill-icon").css("border-bottom", `3px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                updateCellData("color", this.color);
            }
        }
    }
});

//to display colorpicker even on click of image
$("#color-fill-icon", "#text-fill-icon").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10)

})

$(".container").click(function (e) {
    $(".sheet-options-modal").remove();
});

function selectSheet(ele) {

    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    emptySheet();
    selectedSheet = $(ele).text();
    loadSheet();
}

function emptySheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                "font-family": "Noto Sans",
                "font-size": 14,
                "background-color": "#fff",
                "color": "#444",
                "font-weight": "",
                "font-style": "",
                "text-decoration": "",
                "text-align": "left"
            });
        }

    }
}

function loadSheet() { //for loading old sheets
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`); // first cell that have changes
            cell.text(data[rowId][colId].text);
            $("#function-input").text("");
            cell.css({
                "font-family": data[rowId][colId]["font-family"],
                "font-size": data[rowId][colId]["font-size"],
                "background-color": data[rowId][colId]["bgcolor"],
                "color": data[rowId][colId].color,
                "font-weight": data[rowId][colId].bold ? "bold" : "",
                "font-style": data[rowId][colId].italic ? "italic" : "",
                "text-decoration": data[rowId][colId].underlined ? "underline" : "",
                "text-align": data[rowId][colId].alignment
            });
        }
    }
}

$(".addSheet").click(function (e) {
    emptySheet();
    totalSheets++;
    lastlyAddedSheet++;
    while (Object.keys(cellData).includes("Sheet" + lastlyAddedSheet))
        lastlyAddedSheet++;
    selectedSheet = `Sheet${lastlyAddedSheet}`;
    cellData[selectedSheet] = [];

    $(".sheet-tab.selected").removeClass("selected"); //remove selected from previous sheet
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">${selectedSheet}</div>`); //add new sheet tab
    $(".sheet-tab.selected")[0].scrollIntoView();
    addTabClickEvents();
    $(`#row-1-col-1`).click();//to change header 
    saved=false;

})

function addTabClickEvents() {
    //click events work even after removal of selected classes as they get mapped to element itself 
    $(".sheet-tab.selected").bind("contextmenu", function (e) {
        e.preventDefault();

        $(".sheet-options-modal").remove();
        let modal = $(`<div class="sheet-options-modal">
        <div class="option sheet-rename">Rename</div>
        <div class="option sheet-delete">Delete</div>
    </div> `);

        $(".container").append(modal);
        $(".sheet-options-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX })

        $(".sheet-rename").click(function (e) {
            let renameModal = ` <div class="sheet-modal-parent">
                            <div class="sheet-rename-modal">
                                <div class="sheet-modal-title">
                                    <span>Rename Sheet</span>
                                </div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">Rename Sheet to:</span>
                                    <input class="sheet-modal-input" type="text" />
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="btn ok-btn">OK</div>
                                    <div class="btn cancel-btn">Cancel</div>
                                </div>
                            </div>
                        </div>`;
            $(".container").append(renameModal);

            $(".cancel-btn").click(function (e) {
                $(".sheet-modal-parent").remove();
            })

            $(".ok-btn").click(function (e) {
                renameSheet();
            })
        });

        $(".sheet-delete").click(function (e) {
            let deleteModal = `<div class="sheet-modal-parent">
                            <div class="sheet-delete-modal">
                                <div class="sheet-modal-title">
                                    <span>${selectedSheet}</span>
                                </div>
                                <div class="sheet-modal-detail-container">
                                    <span class="sheet-modal-detail-title">Are you sure?</span>
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="btn delete-btn">
                                        <div class="material-icons delete-icon">delete</div>
                                        Delete
                                    </div>
                                    <div class="btn cancel-btn">Cancel</div>
                                </div>
                            </div>
                        </div>`;

            $(".container").append(deleteModal);
            $(".cancel-btn").click(function (e) {
                $(".sheet-modal-parent").remove();
            })
            $(".delete-btn").click(function (e) {
                deleteSheet();
            })
        })

        if (!$(this).hasClass("selected")) //agr aisi old sheet k tab p right click kia h jo selected ni thi to use sath sath select bhi krdo
            selectSheet(this);

    })

    $(".sheet-tab.selected").click(function (e) {
        if (!$(this).hasClass("selected"))
            selectSheet(this);
        $(`#row-1-col-1`).click(); //to change header
    })
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) { //agr already vo name ni exist krta
        let newCellData = {};
        for (let i of Object.keys(cellData)) {
            if (i == selectedSheet)
                newCellData[newSheetName] = cellData[i];
            else
                newCellData[i] = cellData[i];
        }
        cellData = newCellData;
        selectedSheet = newSheetName;
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
        saved=false;
    }
    else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`<div class="error">Please type a valid sheet name</div> `)
    }
}

function deleteSheet() {
    if (totalSheets > 1) {
        $(".sheet-modal-parent").remove();
        let keysArray = Object.keys(cellData);
        let selectedSheetIndex = keysArray.indexOf(selectedSheet);
        let currentSelectedSheet = $(".sheet-tab.selected");

        if (selectedSheetIndex == 0)
            selectSheet(currentSelectedSheet.next()[0]);
        else
            selectSheet(currentSelectedSheet.prev()[0]);

        delete cellData[currentSelectedSheet.text()];
        currentSelectedSheet.remove(); //to remove tab
        totalSheets--;
        saved=false;
    }
}

$(".left-scroller").click(function (e) {
    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != 0)
        selectSheet($(".sheet-tab.selected").prev()[0]);
    $(".sheet-tab.selected")[0].scrollIntoView();
})

$(".right-scroller").click(function (e) {
    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != (keysArray.length - 1)) {
        selectSheet($(".sheet-tab.selected").next()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView();
})

$(`#file-menu`).click(function (e) {
    let fileModal = $(`<div class="file-modal-parent">
                        <div class="options-modal">
                            <div class="close">
                                <div class="material-icons close-icon">arrow_circle_down</div>
                                <div>Close</div>
                            </div>
                            <div class="new">
                                <div class="material-icons new-icon">insert_drive_file</div>
                                <div>New</div>
                            </div>
                            <div class="open">
                                <div class="material-icons open-icon">folder_open</div>
                                <div>Open</div>
                            </div>
                            <div class="save">
                                <div class="material-icons save-icon">save</div>
                                <div>Save</div>
                            </div>
                        </div>
                        <div class="recent-files-modal">
                        <div class="recent">RECENT FILES</div>
                        <div class="allFiles"></div>
                        </div>
                        <div class="transparent-file-modal"></div>
                    </div>`);

    $(`.container`).append(fileModal);
    

    fileModal.animate({
        "width":"100vw"
    },300);
    getData();

    $(".close,.transparent-file-modal,.new,.open,.save").click(function(e){
            fileModal.animate({
               "width":"0vw"
         },300);
         setTimeout(() => {
            fileModal.remove();
        }, 300);
    })

    $(`.new`).click(function(e){
        if(saved){
           newFile();
        }else{
            $(`.container`).append(`<div class="sheet-modal-parent">
                                        <div class="sheet-delete-modal">
                                            <div class="sheet-modal-title">
                                                <span>${$(".title-bar").text()}</span>
                                            </div>
                                            <div class="sheet-modal-detail-container">
                                                <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                                            </div>
                                            <div class="sheet-modal-confirmation">
                                                <div class="btn delete-btn">
                                                    Save
                                                </div>
                                                <div class="btn cancel-btn">Cancel</div>
                                            </div>
                                        </div>
                                    </div>`);

              $(".delete-btn").click(function(e){
                $(".sheet-modal-parent").remove();
                saveFile(true);
              })  
              
              $(".cancel-btn").click(function(e){
                  $(".sheet-modal-parent").remove();
                  newFile();
              })
        }
    })

    $(".save").click(function(e){
        saveFile();
    })

    $(".open").click(function(e){
        openFile();
    })

})

function newFile(){
    emptySheet();
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected>Sheet1</div>`);
    cellData={"Sheet1":{}};
    totalSheets=1;
    lastlyAddedSheet=1;
    selectedSheet="Sheet1";
    addEventsToCells();
    $(`#row-1-col-1`).click();
}

function saveFile(createNewFile){
    $(".container").append(`<div class="sheet-modal-parent">
                                <div class="sheet-rename-modal">
                                    <div class="sheet-modal-title">
                                        <span>Save file as</span>
                                    </div>
                                    <div class="sheet-modal-input-container">
                                        <span class="sheet-modal-input-title">File Name</span>
                                        <input class="sheet-modal-input" type="text" value='${$(".title-bar").text()}'/>
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="btn ok-btn">OK</div>
                                        <div class="btn cancel-btn">Cancel</div>
                                    </div>
                                </div>
                            </div>`);

    $(".cancel-btn").click(function(e){
        $(".sheet-modal-parent").remove();
        if(createNewFile) newFile();
    })   
    
    $(".ok-btn").click(function(e){
        let fileName=$(".sheet-modal-input").val();
        if(fileName){
            let href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
            let a=$(`<a href=${href} download="${fileName}.json"></a>`);
            $(".container").append(a);
            a[0].click();
           
            a.remove();
            saved=true;

            $(".sheet-modal-parent").remove();
            if(createNewFile) newFile();
            addData(fileName,JSON.stringify(cellData));
        }
       
    })
}

function openFile(){
    let inputFile= $(`<input accept="application/json" type="file"/>`);
    $(".container").append(inputFile);
    inputFile.click();
    inputFile.change(function(e){
        let file= e.target.files[0];
        $(".title-bar").text(file.name.split(".json")[0]);
        let reader= new FileReader();
        reader.readAsText(file);
        reader.onload=function(){
            emptySheet(); 
            $(".sheet-tab").remove();
            cellData=JSON.parse(reader.result);
            let sheets= Object.keys(cellData);
            for(let i of sheets)
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`);
            addTabClickEvents();//works on only selected sir
            $(".sheet-tab").removeClass("selected");
            $($(".sheet-tab")[0]).addClass("selected");
            selectedSheet = sheets[0];
            totalSheets =sheets.length;
            lastlyAddedSheet=totalSheets;
            loadSheet();
            
        }
        inputFile.remove();
    })

}

let clipboard={startcell:[],cellData:{}};
let contentCut=false;

$("#cut,#copy").click(function(e){
    if($(this).text() == "content_cut")
        contentCut=true;

   clipboard.startcell=getIDs($(".input-cell.selected")[0]);
   $(".input-cell.selected").each((index,data) => {
       let [rowId,colId]=getIDs(data);
       if(cellData[selectedSheet][rowId-1] && cellData[selectedSheet][rowId-1][colId-1]) //cell m data h
       {
           if(!clipboard.cellData[rowId])
           clipboard.cellData[rowId]={};
       }
       clipboard.cellData[rowId][colId]={...cellData[selectedSheet][rowId-1][colId-1]};      
   });
});

$("#paste").click(function(e){
    if(contentCut) emptySheet();
    let startcell=getIDs($(".input-cell.selected")[0]);
    let rows= Object.keys(clipboard.cellData);
    for(let i of rows){
        let cols= Object.keys(clipboard.cellData[i]);
        for(let j of cols){
            if(contentCut){
                delete cellData[selectedSheet][i-1][j-1];
                if(Object.keys(cellData[selectedSheet][i-1]).length==0)
                  delete cellData[selectedSheet][i-1];
            }
            let rowDistance = parseInt(i) - parseInt(clipboard.startcell[0]);
            let colDistance = parseInt(j) - parseInt(clipboard.startcell[1]);
            if(!cellData[selectedSheet][startcell[0]-rowDistance-1])
               cellData[selectedSheet][startcell[0]-rowDistance-1]={};  
            cellData[selectedSheet][startcell[0] + rowDistance - 1][startcell[1] + colDistance - 1] = {...clipboard.cellData[i][j]};   
        }
    }
    if(contentCut){
        contentCut=false;
        clipboard={startcell:[],cellData:{}};
    }
    loadSheet();
  
})

$("#function-input").blur(function(){
    if($(".input-cell.selected").length >0 ){
        let formula=$(this).text();
            let tempElements=formula.split(" ");
            let elements=[];
            for (let i of tempElements) {
                if (i.length > 1) {
                    i = i.replace("(", "");
                    i = i.replace(")", "");
                    elements.push(i);
                }
            }
            $(".input-cell.selected").each(function(index,data){
                if(updateStreams(data,elements,false)){
                    let [rowId, colId] = getIDs(data);
                    cellData[selectedSheet][rowId - 1][colId - 1].formula = formula;
                    let selfCode = colIdtoName(colId) + rowId;
                    evalFormula(selfCode);
                }
                else alert("Formula invalid");  
            });
        
    }else{
        alert("Cell not selected.");
    }       
});

function updateStreams(ele,elements,update,oldUpstream) {
    let [rowId, colId] = getIDs(ele); //jis cell m formula lgana h uski ids
    let selfCode = colIdtoName(colId) + rowId; //ele n apna name nikaala
    for(let i = 0; i < elements.length; i++) {
        if(checkForSelf(rowId,colId,elements[i])){ //agr ye ele khud m h
            return false;
        }
    }
    //if upstream update krre h firse usi element ki to uska upstream phle se exist krta h or aise h to usme phle to cell th unka downstream bhi update krna pdega
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]){
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
        
        for (let i of downStream) {
            if (elements.includes(i)) {
                return false;
            }
        }

        for (let i of downStream) {
            let [calRowId, calColId] = findIDsFromName(i);
            updateStreams($(`#row-${calRowId}-col-${calColId}`)[0], elements, true, upStream);
        }
    }

    if (!cellData[selectedSheet][rowId - 1]) { //row exist ni krti
        cellData[selectedSheet][rowId - 1] = {};
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else if (!cellData[selectedSheet][rowId - 1][colId - 1]) { //cell exist ni krta
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    }else{
        let upStream= [...cellData[selectedSheet][rowId - 1][colId - 1].upStream];
        if(update){
            for (let i of oldUpstream) {
                let [calRowId, calColId] = findIDsFromName(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfCode);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
                index = cellData[selectedSheet][rowId - 1][colId - 1].upStream.indexOf(i);
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.splice(index, 1);
            }
            for (let i of elements) {
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.push(i);
            }
        }else{
            for(let i of upStream){
                let [calRowId, calColId] = findIDsFromName(i);//ele k upstream k elements ki ids 
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfCode);//un ids k downstream ka vo index jha upstream wala ele h
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);//delete that ele 
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)){
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) 
                        delete cellData[selectedSheet][calRowId - 1];
                }
            }
            cellData[selectedSheet][rowId - 1][colId - 1].upStream = [...elements]; //formula change hua h to purani del krdo
        }

    }
    for (let i = 0; i < elements.length; i++) {
        let [calRowId, calColId] = findIDsFromName(elements[i]);
        if (!cellData[selectedSheet][calRowId - 1]) {
            cellData[selectedSheet][calRowId - 1] = {};
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfCode] };
        } else if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfCode] };
        } else {
            cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfCode);
        }
    }
    return true;
}

function checkForSelf(rowId,colId,ele){
    let [calRowId,calColId]= findIDsFromName(ele);

    if(calRowId==rowId && calColId== colId) //to vo khud formula m h
      return true;
    
   else{
       let selfName = colIdtoName(colId)+rowId; //cell ka naam
       if (!cellData[selectedSheet][calRowId - 1]) {
        cellData[selectedSheet][calRowId - 1] = {};
        cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [] };
       } else if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
        cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [] };
       }

       if (!cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.includes(selfName)) { //calculated cell k downstteam m daalo
        cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfName);
    }

    return false;
   }   
}

function findIDsFromName(ele){ //to separate roewname and col name and find their ids 
    let calRowId;
    let calColId;
    for(let i=0;i<ele.length;i++){
        if(!isNaN(ele[i])) { //no. hai
          let leftString= ele.substring(0,i);
          let rightString=ele.substring(i);
          calRowId=parseInt(rightString);
          calColId = colNameToId(leftString);//calc col id using func name to id
          break;
        }
    } 
    return [calRowId,calColId];      
}

function colNameToId(str){ //opposite of colIdtoName
    let len=str.length;
    let power=len-1;
    let code=0;
    for(let i=0;i<len;i++){
       code= code + (str.charCodeAt(i)-64)* Math.pow( 26,power);
       power--;
    }
    return code;
}


function evalFormula(cell) {
    let [rowId, colId] = findIDsFromName(cell);
    let formula = cellData[selectedSheet][rowId - 1][colId - 1].formula;
    
    if (formula != "") {
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        let upStreamValue = [];
        for (let i in upStream) {
            let [calRowId, calColId] = findIDsFromName(upStream[i]);
            let value;
            if (cellData[selectedSheet][calRowId - 1][calColId - 1].text == "") {
                value = "0";
            }
             else {
                value = cellData[selectedSheet][calRowId - 1][calColId - 1].text;
            }
            upStreamValue.push(value);
            formula = formula.replace(upStream[i], upStreamValue[i]);
        }
        cellData[selectedSheet][rowId - 1][colId - 1].text = eval(formula);
        loadSheet();
    }
    let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
    for (let i = downStream.length - 1; i >= 0; i--) {
        evalFormula(downStream[i]);
    }

}




// function addLoader(){
//     $(".container").append(`<div class="sheet-modal-parent loader-parent">
//                             <div class="loading-img"><img src="loader.gif"/></div>
//                             <div class="loading">Loading...</div>
//                             </div>`)
// }  

// function removeLoader(){
//     $(".loader-parent").remove();

