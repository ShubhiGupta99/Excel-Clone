let request= indexedDB.open("fileDB",1); //opens if does not exist else creates

let db;
request.onsuccess = function(e){
    db=request.result; 
}
request.onupgradeneeded=function(e){
    db=request.result;
    db.createObjectStore("fileStore",{keyPath:"nId"})  //creating obj or table gallery with primary key nid
}


function addData(name,data){
    let tx = db.transaction("fileStore","readwrite");
    let store= tx.objectStore("fileStore"); //choose table
    store.add({nId:Date.now(),fileName:name,data:data});
}



function getData(){
    let tx = db.transaction("fileStore","readonly");
    let store =tx.objectStore("fileStore");
    let req= store.openCursor();
    $(".allFiles").innerHTML="";
    req.onsuccess = function (e) {
        let cursor=req.result;
        if(cursor){
            let d = new Date(cursor.value.nId);
            d=d.toString().split(" ");
            let date=" ";
            console.log(d);
            for(let i=1;i<5;i++) date=date+d[i]+" ";
            
           let fileDiv = `<div class="file"><img style="margin-right:15px"; src="https://img.icons8.com/fluent/22/000000/microsoft-excel-2019.png"/>
           <div>${cursor.value.fileName}</div>
           <div style="font-size:15px;margin-left:20px;">Last modified on: ${date}</div> 
           </div>`;
           $(".allFiles").append(fileDiv);
           cursor.continue();
        } 
    }
}    