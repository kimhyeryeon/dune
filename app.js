/* ==========================
   Dune Wiki
   app.js
========================== */

let dictionary = [];
let filtered = [];

/* 시작 */
window.onload = function () {

    loadExcel();

    document
        .getElementById("searchInput")
        .addEventListener("input", search);

};


/* 엑셀 읽기 */
async function loadExcel() {

    const response = await fetch("dune_dictionary.xlsx");

    const data = await response.arrayBuffer();

    const workbook = XLSX.read(data, { type: "array" });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    dictionary = XLSX.utils.sheet_to_json(sheet);

    filtered = dictionary;

    drawList();

}


/* 목록 출력 */

function drawList() {

    const list = document.getElementById("termList");

    list.innerHTML = "";

    document.getElementById("resultCount").textContent =
        filtered.length + "개의 용어";

    filtered.forEach(item => {

        const div = document.createElement("div");

        div.className = "term";

        div.textContent = item["용어 (원어)"];

        div.onclick = function () {

            show(item);

        };

        list.appendChild(div);

    });

}


/* 검색 */

function search() {

    const keyword =
        document
        .getElementById("searchInput")
        .value
        .trim()
        .toLowerCase();

    filtered = dictionary.filter(item => {

        return (

            String(item["용어 (원어)"])
            .toLowerCase()
            .includes(keyword)

            ||

            String(item["정의 및 설명"])
            .toLowerCase()
            .includes(keyword)

            ||

            String(item["페이지"])
            .includes(keyword)

        );

    });

    drawList();

}


/* 문서 보기 */

function show(item) {

    document.getElementById("welcome").style.display = "none";

    document.getElementById("article").style.display = "block";

    document.getElementById("termTitle").textContent =
        item["용어 (원어)"];

    document.getElementById("termOrigin").textContent =
        "페이지 : " + item["페이지"];

    document.getElementById("termPage").textContent = "";

    document.getElementById("termDescription").textContent =
        item["정의 및 설명"];

}