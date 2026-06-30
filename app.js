/* =====================================
   Dune Wiki Mobile
===================================== */

let dictionary = [];
let filtered = [];

/* -------------------------------
   시작
-------------------------------- */

window.onload = function () {

    document
        .getElementById("searchInput")
        .addEventListener("input", searchWord);

    document
        .getElementById("backButton")
        .addEventListener("click", showList);

    loadExcel();

};


/* -------------------------------
   엑셀 읽기
-------------------------------- */

async function loadExcel() {

    try {

        const response = await fetch("dune_dictionary.xlsx");

        const data = await response.arrayBuffer();

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const sheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];

        dictionary =
            XLSX.utils.sheet_to_json(sheet);

        filtered = [...dictionary];

        drawList();

    }

    catch (e) {

        alert("엑셀을 읽을 수 없습니다.");

        console.error(e);

    }

}


/* -------------------------------
   목록 출력
-------------------------------- */

function drawList() {

    const list =
        document.getElementById("termList");

    list.innerHTML = "";

    document
        .getElementById("resultCount")
        .textContent =
        filtered.length + "개의 용어";

    filtered.forEach(item => {

        const div =
            document.createElement("div");

        div.className = "term";

        div.textContent =
            item["용어 (원어)"];

        div.onclick = function () {

            showArticle(item);

        };

        list.appendChild(div);

    });

}


/* -------------------------------
   검색
-------------------------------- */

function searchWord() {

    const keyword =
        document
            .getElementById("searchInput")
            .value
            .trim()
            .toLowerCase();

    if (keyword === "") {

        filtered = [...dictionary];

        drawList();

        return;

    }

    filtered =
        dictionary.filter(item => {

            const title =
                String(
                    item["용어 (원어)"] || ""
                ).toLowerCase();

            const desc =
                String(
                    item["정의 및 설명"] || ""
                ).toLowerCase();

            const page =
                String(
                    item["페이지"] || ""
                );

            return (

                title.includes(keyword)

                ||

                desc.includes(keyword)

                ||

                page.includes(keyword)

            );

        });

    drawList();

}


/* -------------------------------
   설명 보기
-------------------------------- */

function showArticle(item) {

    document
        .getElementById("listPage")
        .style.display = "none";

    document
        .getElementById("articlePage")
        .style.display = "flex";

    document
        .getElementById("termTitle")
        .textContent =
        item["용어 (원어)"];

    document
        .getElementById("termOrigin")
        .textContent =
        "";

    document
        .getElementById("termPage")
        .textContent =
        "페이지 : " +
        item["페이지"];

    document
        .getElementById("termDescription")
        .textContent =
        item["정의 및 설명"];

}


/* -------------------------------
   목록으로
-------------------------------- */

function showList() {

    document
        .getElementById("articlePage")
        .style.display = "none";

    document
        .getElementById("listPage")
        .style.display = "flex";

}
