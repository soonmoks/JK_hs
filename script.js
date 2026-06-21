let admissionData = [];
let selectedCollege = "";
let currentCollegeRows = [];
let currentMajorName = "";
let currentSimilarFilter = "all";
let currentMajorKeyword = "";
let currentAdmissionRows = [];

/* =========================
   메뉴 전환
========================= */

function showPage(pageId, button){

    document
    .querySelectorAll(".page")
    .forEach(page=>{
        page.classList.add("hidden");
    });

    document
    .getElementById(pageId)
    .classList.remove("hidden");

    document
    .querySelectorAll(".sidebar button")
    .forEach(btn=>{
        btn.classList.remove("active");
    });

    button.classList.add("active");
}

/* =========================
   엑셀 업로드
========================= */

document
.getElementById("excelFile")
.addEventListener("change", loadExcel);

function loadExcel(event){

    const file = event.target.files[0];

    if(!file) return;

    const reader = new FileReader();

    reader.onload = function(e){

        const data =
        new Uint8Array(e.target.result);

        const workbook =
        XLSX.read(data,{
            type:"array"
        });

        const sheetName =
        workbook.SheetNames[0];

        const worksheet =
        workbook.Sheets[sheetName];

        admissionData =
        XLSX.utils.sheet_to_json(worksheet);

        buildFilters();

        updateDashboard();

        initAdmissionFilters();

        initRegionAnalysis();

        buildCollegeSearch();

        buildMajorSearch();

        document
        .getElementById("uploadInfo")
        .innerHTML =
        `
        파일명 : ${file.name}<br>
        데이터 건수 : ${admissionData.length.toLocaleString()}
        `;
    };

    reader.readAsArrayBuffer(file);
}

/* =========================
   필터 생성
========================= */

function buildFilters(){

    buildFilter(
        "yearFilter",
        "학년도"
    );

    buildFilter(
        "regionFilter",
        "지역"
    );

    buildFilter(
        "periodFilter",
        "모집시기"
    );

    document
    .getElementById("yearFilter")
    .addEventListener("change",applyFilters);

    document
    .getElementById("regionFilter")
    .addEventListener("change",applyFilters);

    document
    .getElementById("periodFilter")
    .addEventListener("change",applyFilters);

    document
.getElementById("scoreFilter")
.addEventListener("change",applyFilters);
}

function buildFilter(selectId,columnName){

    const select =
    document.getElementById(selectId);

    select.innerHTML =
    `<option value="">전체</option>`;

    const values =
    [...new Set(
        admissionData
        .map(r=>r[columnName])
        .filter(v=>v)
    )]
    .sort();

    values.forEach(value=>{

        const option =
        document.createElement("option");

        option.value = value;
        option.textContent = value;

        select.appendChild(option);

    });
}

/* =========================
   필터 적용
========================= */

function getFilteredData(){

    let data =
    [...admissionData];

    const year =
    document.getElementById("yearFilter").value;

    const region =
    document.getElementById("regionFilter").value;

    const period =
    document.getElementById("periodFilter").value;

    if(year){

        data =
        data.filter(row=>
            String(row["학년도"]) === String(year)
        );
    }

    if(region){

        data =
        data.filter(row=>
            row["지역"] === region
        );
    }

    if(period){

        data =
        data.filter(row=>
            row["모집시기"] === period
        );
    }

    return data;
}

function applyFilters(){

    updateDashboard();
    analyzeAdmission();

    if(selectedCollege){

        analyzeCollege(
            selectedCollege
        );

        if(currentMajorName){

            analyzeMajor(
                currentMajorName,
                currentCollegeRows
            );

            const score =
            Number(
                document
                .getElementById(
                    "myScoreInput"
                ).value
            );

            if(!isNaN(score)){

                analyzePrediction();

            }
        }
    }

    if(currentMajorKeyword){

        analyzeMajorGlobal(
            currentMajorKeyword
        );

    }

    analyzeRegion();
}

/* =========================
   대시보드
========================= */

function updateDashboard(){

    const filteredData =
    getFilteredData();

    const studentSet =
    new Set();

    const collegeSet =
    new Set();

    const majorSet =
    new Set();

    let passCount = 0;
    let failCount = 0;

    filteredData.forEach(row=>{

        if(row["이름"]){
            studentSet.add(row["이름"]);
        }

        if(row["대학명"]){
            collegeSet.add(row["대학명"]);
        }

        if(row["모집단위"]){
            majorSet.add(row["모집단위"]);
        }

        if(row["최종단계"]==="합격"){
            passCount++;
        }else{
            failCount++;
        }

    });

    const applyCount =
    filteredData.length;

    const passRate =
    applyCount===0
    ? 0
    : ((passCount/applyCount)*100)
      .toFixed(2);

    document.getElementById("studentCount").innerText =
    studentSet.size.toLocaleString();

    document.getElementById("applyCount").innerText =
    applyCount.toLocaleString();

    document.getElementById("passCount").innerText =
    passCount.toLocaleString();

    document.getElementById("failCount").innerText =
    failCount.toLocaleString();

    document.getElementById("passRate").innerText =
    passRate + "%";

    document.getElementById("collegeCount").innerText =
    collegeSet.size.toLocaleString();

    document.getElementById("majorCount").innerText =
    majorSet.size.toLocaleString();
}

/* =========================
   대학 자동완성
========================= */

function buildCollegeSearch(){

    const searchBox =
    document.getElementById("collegeSearch");

    if(!searchBox) return;

    searchBox.oninput = function(){

        const keyword =
        this.value.trim();

        const suggestionBox =
        document.getElementById("collegeSuggestions");

        suggestionBox.innerHTML = "";

        if(keyword.length < 1){
            return;
        }

        const collegeList =
        [...new Set(
            admissionData
            .map(r=>r["대학명"])
            .filter(v=>v)
        )];

        const matched =
        collegeList.filter(college=>
            college.includes(keyword)
        );

        matched
        .slice(0,20)
        .forEach(college=>{

            const div =
            document.createElement("div");

            div.className =
            "suggestion-item";

            div.innerText =
            college;

            div.onclick = ()=>{

                searchBox.value =
                college;

                suggestionBox.innerHTML =
                "";

                selectedCollege =
                college;

                analyzeCollege(college);
            };

            suggestionBox.appendChild(div);
        });
    };
}

/* =========================
   대학 분석
========================= */

function analyzeCollege(collegeName){

    const filteredData =
    getFilteredData();

    const rows =
    filteredData.filter(row=>
        row["대학명"]===collegeName
    );

    currentCollegeRows = rows;

document
.getElementById("selectedCollegeTitle")
.innerText = collegeName;

buildStudentTables(rows);

    let passCount = 0;
    let failCount = 0;

    rows.forEach(row=>{

        if(row["최종단계"]==="합격"){
            passCount++;
        }else{
            failCount++;
        }

    });

    const applyCount =
    rows.length;

    const passRate =
    applyCount===0
    ? 0
    : ((passCount/applyCount)*100)
      .toFixed(2);

    document
    .getElementById("collegeApplyCount")
    .innerText =
    applyCount.toLocaleString();

    document
    .getElementById("collegePassCount")
    .innerText =
    passCount.toLocaleString();

    document
    .getElementById("collegeFailCount")
    .innerText =
    failCount.toLocaleString();

    document
    .getElementById("collegePassRate")
    .innerText =
    passRate + "%";

    updateScoreCards(rows);
    buildMajorTable(rows);
}

/* =========================
   모집단위 테이블
========================= */

function buildMajorTable(rows){

    const tbody =
    document.getElementById(
        "majorTableBody"
    );

    tbody.innerHTML = "";

    const majorMap = {};

    rows.forEach(row=>{

        const key =
            row["대학명"] +
            "||" +
            row["모집단위"];

        if(!majorMap[key]){

            majorMap[key] = {

                college :
                row["대학명"],

                major :
                row["모집단위"],

                apply : 0,

                pass : 0,

                fail : 0

            };
        }

        majorMap[key].apply++;

        if(
            row["최종단계"]==="합격"
        ){

            majorMap[key].pass++;

        }else{

            majorMap[key].fail++;

        }

    });

    Object.values(majorMap)

    .sort((a,b)=>

        b.apply - a.apply

    )

    .forEach(item=>{

        const rate =

        item.apply

        ?

        (
            item.pass /
            item.apply *
            100
        ).toFixed(1)

        :

        "0.0";

        const tr =
        document.createElement("tr");

        tr.style.cursor =
        "pointer";

        tr.onclick = ()=>{

            analyzeMajor(
                item.major,
                rows
            );

        };

        tr.innerHTML = `

            <td>${item.college}</td>

            <td>${item.major}</td>

            <td>${item.apply}</td>

            <td>${item.pass}</td>

            <td>${item.fail}</td>

            <td>${rate}%</td>

        `;

        tbody.appendChild(tr);

    });

}

/* =========================
   학생 테이블
========================= */

function buildStudentTables(rows){

    const scoreColumn =
    document
    .getElementById("scoreFilter")
    .value;

    const passBody =
    document
    .getElementById("passStudentBody");

    const failBody =
    document
    .getElementById("failStudentBody");

    passBody.innerHTML = "";
    failBody.innerHTML = "";

    rows.forEach(row=>{

        const score =
        row[scoreColumn] ?? "";

        const tr =
        document.createElement("tr");

        tr.className =
        "student-row";

       const studentCode =
row["학년도"] +
"_" +
String(row["학년"]) +
String(row["반"]).padStart(2,"0") +
String(row["번호"]).padStart(2,"0");

tr.innerHTML = `
<td>${row["이름"]} (${studentCode})</td>
<td>${row["학년도"]}</td>
<td>${row["모집단위"]}</td>
<td>${row["전형명(대)"] || ""}</td>
<td>${score}</td>
`;

        tr.onclick = ()=>{

            showStudentDetail(row);

        };

        if(row["최종단계"]==="합격"){

            passBody.appendChild(tr);

        }else{

            failBody.appendChild(tr);

        }

    });
}

/* =========================
   학생 상세 팝업
========================= */

function showStudentDetail(row){

    const key =
    `${row["학년도"]}-${row["학년"]}-${row["반"]}-${row["번호"]}-${row["이름"]}`;

    const studentRows =
    admissionData.filter(r=>{

        return (
            r["학년도"]==row["학년도"] &&
            r["학년"]==row["학년"] &&
            r["반"]==row["반"] &&
            r["번호"]==row["번호"] &&
            r["이름"]==row["이름"]
        );

    });

    const scoreColumn =
    document
    .getElementById("scoreFilter")
    .value;

    const passCount =
    studentRows.filter(
        r=>r["최종단계"]==="합격"
    ).length;

    const failCount =
    studentRows.length - passCount;

    document
    .getElementById("studentModal")
    .style.display = "flex";



    const studentCode =
row["학년도"] +
"_" +
String(row["학년"]) +
String(row["반"]).padStart(2,"0") +
String(row["번호"]).padStart(2,"0");

document
.getElementById("studentModalTitle")
.innerText =
`${row["이름"]} (${studentCode})`;



    document
    .getElementById("studentSummary")
    .innerHTML = `
        <div class="summary-card">
            <div class="summary-title">
                학년도
            </div>
            <div class="summary-value">
                ${row["학년도"]}
            </div>
        </div>

        <div class="summary-card">
            <div class="summary-title">
                성적
            </div>
            <div class="summary-value">
                ${row[scoreColumn] ?? ""}
            </div>
        </div>

        <div class="summary-card">
            <div class="summary-title">
                지원건수
            </div>
            <div class="summary-value">
                ${studentRows.length}
            </div>
        </div>

        <div class="summary-card">
            <div class="summary-title">
                합격건수
            </div>
            <div class="summary-value">
                ${passCount}
            </div>
        </div>

        <div class="summary-card">
            <div class="summary-title">
                불합격건수
            </div>
            <div class="summary-value">
                ${failCount}
            </div>
        </div>
    `;

    const tbody =
    document
    .getElementById("studentHistoryBody");

    tbody.innerHTML = "";

    studentRows.forEach(r=>{

        const tr =
        document.createElement("tr");

        tr.innerHTML = `
            <td>${r["대학명"]||""}</td>
            <td>${r["모집단위"]||""}</td>
            <td>${r["전형명(대)"]||""}</td>
            <td>${r["최종단계"]||""}</td>
        `;

        tbody.appendChild(tr);

    });

}

/* =========================
   팝업 닫기
========================= */

document
.getElementById("closeModalBtn")
.addEventListener("click",()=>{

    document
    .getElementById("studentModal")
    .style.display = "none";

});

/* =========================
   탭
========================= */

document
.getElementById("tabStat")
.addEventListener("click",()=>{

    showCollegeTab("stat");

});

document
.getElementById("tabPass")
.addEventListener("click",()=>{

    showCollegeTab("pass");

});

document
.getElementById("tabFail")
.addEventListener("click",()=>{

    showCollegeTab("fail");

});


document
.getElementById("tabMajorDetail")
.addEventListener("click",()=>{

    showCollegeTab("major");

});


function showCollegeTab(type){

    document
    .querySelectorAll(".tab-button")
    .forEach(btn=>
        btn.classList.remove("active-tab")
    );

    document
    .getElementById("tabStatContent")
    .style.display = "none";

    document
    .getElementById("tabPassContent")
    .style.display = "none";

    document
    .getElementById("tabFailContent")
    .style.display = "none";

    document
.getElementById("tabMajorDetailContent")
.style.display = "none";


    if(type==="stat"){

        document
        .getElementById("tabStat")
        .classList.add("active-tab");

        document
        .getElementById("tabStatContent")
        .style.display = "block";
    }

    if(type==="pass"){

        document
        .getElementById("tabPass")
        .classList.add("active-tab");

        document
        .getElementById("tabPassContent")
        .style.display = "block";
    }

    if(type==="fail"){

        document
        .getElementById("tabFail")
        .classList.add("active-tab");

        document
        .getElementById("tabFailContent")
        .style.display = "block";
    }

    if(type==="major"){

    document
    .getElementById("tabMajorDetail")
    .classList.add("active-tab");

    document
    .getElementById("tabMajorDetailContent")
    .style.display = "block";
}


}


/* =========================
   성적 통계 카드
========================= */

function updateScoreCards(rows){

    const scoreColumn =
    document
    .getElementById("scoreFilter")
    .value;

    const passScores =
    rows
    .filter(r =>
        r["최종단계"]==="합격"
    )
    .map(r =>
        Number(r[scoreColumn])
    )
    .filter(v =>
        !isNaN(v)
    );

    const failScores =
    rows
    .filter(r =>
        r["최종단계"]!=="합격"
    )
    .map(r =>
        Number(r[scoreColumn])
    )
    .filter(v =>
        !isNaN(v)
    );

    const avgPass =
    passScores.length
    ? (
        passScores.reduce(
            (a,b)=>a+b,0
        ) / passScores.length
      ).toFixed(2)
    : "-";

    const avgFail =
    failScores.length
    ? (
        failScores.reduce(
            (a,b)=>a+b,0
        ) / failScores.length
      ).toFixed(2)
    : "-";

    const minPass =
    passScores.length
    ? Math.max(...passScores).toFixed(2)
    : "-";

    const maxFail =
    failScores.length
    ? Math.min(...failScores).toFixed(2)
    : "-";

    document
    .getElementById("avgPassScore")
    .innerText = avgPass;

    document
    .getElementById("avgFailScore")
    .innerText = avgFail;

    document
    .getElementById("minPassScore")
    .innerText = minPass;

    document
    .getElementById("maxFailScore")
    .innerText = maxFail;
}

/* =========================
   모집단위 상세분석
========================= */

function analyzeMajor(
    majorName,
    collegeRows
){

    const scoreColumn =
    document
    .getElementById("scoreFilter")
    .value;

    const rows =
    collegeRows.filter(
        r => r["모집단위"] === majorName
    );

    currentMajorRows = rows;

    currentMajorName = majorName;

    document
    .getElementById("majorDetailArea")
    .style.display = "block";
    showCollegeTab("major");



    document
    .getElementById("majorDetailTitle")
    .innerText =
  `${selectedCollege}
 > ${majorName}`

    const passRows =
    rows.filter(
        r=>r["최종단계"]==="합격"
    );

    const failRows =
    rows.filter(
        r=>r["최종단계"]!=="합격"
    );

    const passScores =
    passRows
    .map(r=>Number(r[scoreColumn]))
    .filter(v=>!isNaN(v));

    const failScores =
    failRows
    .map(r=>Number(r[scoreColumn]))
    .filter(v=>!isNaN(v));

    const applyCount =
    rows.length;

    const passCount =
    passRows.length;

    const failCount =
    failRows.length;

    const passRate =
    applyCount
    ? (
        passCount /
        applyCount * 100
      ).toFixed(1)
    : 0;

    const avgPass =
    passScores.length
    ? (
        passScores.reduce(
            (a,b)=>a+b,0
        ) / passScores.length
      ).toFixed(2)
    : "-";

    const avgFail =
    failScores.length
    ? (
        failScores.reduce(
            (a,b)=>a+b,0
        ) / failScores.length
      ).toFixed(2)
    : "-";

    const cutLine =
    passScores.length
    ? Math.max(...passScores)
      .toFixed(2)
    : "-";

    const failLine =
    failScores.length
    ? Math.min(...failScores)
      .toFixed(2)
    : "-";

    currentMajorCutLine =
Number(cutLine);

currentMajorFailLine =
Number(failLine);

    document
    .getElementById("majorApplyCount")
    .innerText =
    applyCount;

    document
    .getElementById("majorPassCount")
    .innerText =
    passCount;

    document
    .getElementById("majorFailCount")
    .innerText =
    failCount;

    document
    .getElementById("majorPassRate")
    .innerText =
    passRate + "%";

    document
    .getElementById("majorAvgPass")
    .innerText =
    avgPass;

    document
    .getElementById("majorAvgFail")
    .innerText =
    avgFail;

    document
    .getElementById("majorCutLine")
    .innerText =
    cutLine;

    document
    .getElementById("majorFailLine")
    .innerText =
    failLine;

    buildMajorTrend(
        rows,
        scoreColumn
    );

    buildMajorCompare(
    collegeRows,
    scoreColumn
);


}

/* =========================
   모집단위 연도별 추이
========================= */

function buildMajorTrend(
    rows,
    scoreColumn
){

    const tbody =
    document
    .getElementById("majorTrendBody");

    tbody.innerHTML = "";

    const yearMap = {};

    rows.forEach(row=>{

        const year =
        row["학년도"];

        if(!yearMap[year]){

            yearMap[year] = [];
        }

        yearMap[year]
        .push(row);

    });

    Object
    .keys(yearMap)
    .sort()
    .forEach(year=>{

        const yearRows =
        yearMap[year];

        const apply =
        yearRows.length;

        const pass =
        yearRows.filter(
            r=>r["최종단계"]==="합격"
        );

        const passRate =
        apply
        ? (
            pass.length /
            apply * 100
          ).toFixed(1)
        : 0;

        const passScores =
        pass
        .map(r=>
            Number(
                r[scoreColumn]
            )
        )
        .filter(v=>!isNaN(v));

        const cutLine =
        passScores.length
        ? Math.max(...passScores)
          .toFixed(2)
        : "-";

        const tr =
        document
        .createElement("tr");

        tr.innerHTML = `
            <td>${year}</td>
            <td>${apply}</td>
            <td>${pass.length}</td>
            <td>${passRate}%</td>
            <td>${cutLine}</td>
        `;

        tbody.appendChild(tr);

    });

}

/* =========================
   진학 가능성 분석
========================= */

let currentMajorCutLine = null;
let currentMajorFailLine = null;

let currentMajorRows = [];

/* =========================
   진학 가능성 분석
========================= */

document
.getElementById("predictBtn")
.addEventListener(
    "click",
    analyzePrediction
);

function analyzePrediction(){

    const score =
    Number(
        document
        .getElementById("myScoreInput")
        .value
    );

    window.currentMyScore =
score;

    if(isNaN(score)){

        alert("성적을 입력하세요.");

        return;
    }

    let result = "";

    if(score <= currentMajorCutLine - 0.20){

        result = "매우안정";

    }
    else if(score <= currentMajorCutLine){

        result = "안정";

    }
    else if(score <= currentMajorFailLine){

        result = "적정";

    }
    else{

        result = "도전";

    }

    window.currentPredictionResult =
    result;

    document
    .getElementById("predictResult")
    .innerHTML = `
       <div class="card predict-card">
    <h3>판정 결과</h3>
            <span>${result}</span>
        </div>
    `;

    buildSimilarStudents(score);
    updateScorePosition(score);
    buildRecommendMajors(score);
}

/* =========================
   유사 선배 TOP10
========================= */

function buildSimilarStudents(score){

    const scoreColumn =
    document
    .getElementById("scoreFilter")
    .value;

    const tbody =
    document
    .getElementById("similarStudentBody");

    tbody.innerHTML = "";

    const studentMap = {};

    currentMajorRows.forEach(row=>{

    const key =

        row["학년도"] + "|" +
        row["학년"] + "|" +
        row["반"] + "|" +
        row["번호"] + "|" +
        row["이름"];

    if(!studentMap[key]){

        studentMap[key] = row;
    }

    });

    let list =
Object.values(studentMap)
.filter(row=>{

    const s =
    Number(
        row[scoreColumn]
    );

    return !isNaN(s);

})
.map(row=>{

    const s =
    Number(
        row[scoreColumn]
    );

    return {

        row,

        score:s,

        diff:
        Math.abs(
            score - s
        )

    };

});

if(currentSimilarFilter==="pass"){

    list =
    list.filter(item=>

        item.row["최종단계"]==="합격"

    );

}
else if(
    currentSimilarFilter==="fail"
){

    list =
    list.filter(item=>

        item.row["최종단계"]!=="합격"

    );

}

list =

list
.sort(
    (a,b)=>
    a.diff - b.diff
)
.slice(0,10);

    list.forEach(item=>{

        const row =
        item.row;

        const studentRows =
        admissionData.filter(r=>

            r["학년도"]==row["학년도"] &&
            r["학년"]==row["학년"] &&
            r["반"]==row["반"] &&
            r["번호"]==row["번호"] &&
            r["이름"]==row["이름"]

        );

        const studentCode =

            row["학년도"] +
            "_" +

            String(row["학년"]) +

            String(row["반"])
            .padStart(2,"0") +

            String(row["번호"])
            .padStart(2,"0");

        const passCount =
            studentRows.filter(
             r=>r["최종단계"]==="합격"
        ).length;

        const passRate =
    (
    passCount /
    studentRows.length *
    100
        ).toFixed(0);

        const tr =
document
.createElement("tr");

        tr.style.cursor =
        "pointer";

        tr.innerHTML = `
            <td>
    <div style="font-size:13px;color:#5f6368;">
        ${studentCode}
    </div>

    <div style="
        font-size:16px;
        font-weight:600;
        margin-top:4px;">
        ${row["이름"]}
    </div>
</td>

<td>
    ${item.score}
    <br>
    <small class="${
        item.score > score
        ? 'diff-bad'
        : 'diff-good'
    }">
        ${
            item.score > score
            ? '🔴 +' + item.diff.toFixed(2)
            : '🟢 -' + item.diff.toFixed(2)
        }
    </small>
</td>

            <td>

<span class="${
    row["최종단계"]==="합격"
    ? "result-pass"
    : "result-fail"
}">

${row["최종단계"]}

</span>

</td>

<td>

    지원 ${studentRows.length}건

    <br>

    합격 ${passCount}건

    <br>

    <strong style="color:#1a73e8;">
        ${passRate}%
    </strong>

</td>

        `;

        tr.onclick = ()=>{

            showStudentDetail(
                row
            );

        };

        tbody.appendChild(tr);

    });

}

function updateScorePosition(myScore){

    const passAvg =
    Number(
        document
        .getElementById("majorAvgPass")
        .innerText
    );

    const failAvg =
    Number(
        document
        .getElementById("majorAvgFail")
        .innerText
    );

    const cutLine =
    Number(
        document
        .getElementById("majorCutLine")
        .innerText
    );

    document.getElementById("mineScoreText")
    .innerText =
    " " + myScore.toFixed(2);

    document.getElementById("passScoreText")
    .innerText =
    " " + passAvg.toFixed(2);

    document.getElementById("cutScoreText")
    .innerText =
    " " + cutLine.toFixed(2);

    document.getElementById("failScoreText")
    .innerText =
    " " + failAvg.toFixed(2);

    const positions = [

        {
            id:"minePosition",
            label:".mine-label",
            score:myScore
        },

        {
            id:"passPosition",
            label:".pass-label",
            score:passAvg
        },

        {
            id:"cutPosition",
            label:".cut-label",
            score:cutLine
        },

        {
            id:"failPosition",
            label:".fail-label",
            score:failAvg
        }

    ];

    positions.forEach(item=>{

        const y =
        scoreToPosition(item.score);

        document
        .getElementById(item.id)
        .style.top =
        y + "px";

        item.y = y;
    });

    positions.sort(
        (a,b)=>a.y-b.y
    );

    let lastLabelY = -999;

    positions.forEach(item=>{

        let labelY = item.y;

        if(labelY-lastLabelY < 30){

            labelY =
            lastLabelY + 30;
        }

        const label =
        document.querySelector(item.label);

        label.style.top =
        (labelY-item.y) + "px";

        const line =
        label.querySelector(".position-line");

        const dx = 120;
        const dy = labelY-item.y;

        const length =
        Math.sqrt(
            dx*dx + dy*dy
        );

        const angle =
        Math.atan2(
            dy,
            dx
        ) * 180 / Math.PI;

        line.style.width =
        length + "px";

        line.style.transform =
        `rotate(${angle}deg)`;

        lastLabelY = labelY;
    });

    document
    .getElementById("scorePositionArea")
    .classList
    .remove("hidden");

    let predict = 50;

    if(myScore < passAvg){

        predict =
        Math.min(
            95,
            50 +
            ((passAvg-myScore)*40)
        );

    }else{

        predict =
        Math.max(
            5,
            50 -
            ((myScore-passAvg)*40)
        );
    }

    document
    .getElementById("positionPredict")
    .innerText =
    predict.toFixed(0) + "%";

    const rank =
    Math.max(
        1,
        Math.min(
            99,
            100-predict
        )
    );

    document
    .getElementById("positionRank")
    .innerText =
    "TOP " +
    rank.toFixed(0) +
    "%";

    document
    .getElementById("positionResult")
    .innerText =
    window.currentPredictionResult;
}

function updateLabelLine(
    markerId,
    labelSelector
){

    const marker =
    document.getElementById(markerId);

    const label =
    document.querySelector(labelSelector);

    if(!marker || !label) return;

    const markerY =
    marker.offsetTop;

    const labelY =
    marker.offsetTop +
    label.offsetTop;

    const diff =
    labelY - markerY;

    const angle =
    Math.atan2(
        diff,
        140
    ) * 180 / Math.PI;

    label.style.setProperty(
        "--line-angle",
        angle + "deg"
    );
}

function buildRecommendMajors(myScore){

const area =
document.getElementById(
    "recommendArea"
);

area.classList.remove(
    "hidden"
);

const safe = [];
const stable = [];
const normal = [];
const challenge = [];

const scoreField =
document
.getElementById("scoreFilter")
.value;

const filteredData =
getFilteredData();

const majorMap = {};

filteredData.forEach(row=>{

    const key =

        row["대학명"] +
        "||" +
        row["모집단위"];

    if(!majorMap[key]){

        majorMap[key] = [];
    }

    majorMap[key].push(row);

});

Object.entries(majorMap)
.forEach(([key,rows])=>{

    const passRows =
    rows.filter(
        r=>r["최종단계"]==="합격"
    );

    if(passRows.length===0)
        return;

    const scores =
    passRows
    .map(r=>Number(r[scoreField]))
    .filter(v=>!isNaN(v));

    if(scores.length===0)
        return;

    const avgScore =

        scores.reduce(
            (a,b)=>a+b,
            0
        ) / scores.length;

    const diff =
    avgScore - myScore;

    const parts =
    key.split("||");

    const item = {

        college:
        parts[0],

        major:
        parts[1],

        avg:
        avgScore,

        diff:
        diff

    };

    if(diff >= 0.2){

        safe.push(item);

    }
    else if(diff >= 0.1){

        stable.push(item);

    }
    else if(Math.abs(diff)<=0.1){

        normal.push(item);

    }
    else{

        challenge.push(item);

    }

});

safe.sort(
    (a,b)=>a.diff-b.diff
);

stable.sort(
    (a,b)=>a.diff-b.diff
);

normal.sort(
    (a,b)=>a.diff-b.diff
);

challenge.sort(
    (a,b)=>b.diff-a.diff
);

renderRecommend(
    "safeList",
    "🟢 매우 안정",
    safe
);

renderRecommend(
    "stableList",
    "🟡 안정",
    stable
);

renderRecommend(
    "normalList",
    "🟠 적정",
    normal
);

renderRecommend(
    "challengeList",
    "🔴 도전",
    challenge
);

}


function renderRecommend(
targetId,
title,
list
){

const el =
document.getElementById(
    targetId
);

el.innerHTML =
`
<div class="recommend-group">

    <h3>
        ${title}
        (${list.length})
    </h3>

    ${
        list
        .slice(0,20)
        .map(item=>`

            <div class="recommend-major">

                <div class="recommend-college">

                    ${item.college}

                </div>

                <div class="recommend-major-name">

                    ${item.major}

                </div>

                <div class="recommend-extra">

                    합격평균
                    ${item.avg.toFixed(2)}

                </div>

                <div class="recommend-extra
                ${
                    item.diff >= 0
                    ? "diff-good"
                    : "diff-bad"
                }">

                    ${
                        item.diff >= 0
                        ? "+"
                        : ""
                    }

                    ${item.diff.toFixed(2)}

                </div>

            </div>

        `)
        .join("")
    }

</div>
`;

}


function scoreToPosition(score){

    const min = 1;
    const max = 9;

    const ratio =
    (score-min)/(max-min);

    return 40 + (ratio*320);

}

document
.getElementById("minePosition")
.className =
"position-marker minePosition";

document
.getElementById("passPosition")
.className =
"position-marker passPosition";

document
.getElementById("cutPosition")
.className =
"position-marker cutPosition";

document
.getElementById("failPosition")
.className =
"position-marker failPosition";

function buildMajorCompare(
    collegeRows,
    scoreColumn
){

    const tbody =
    document.getElementById(
        "majorCompareBody"
    );

    tbody.innerHTML = "";

    const majorMap = {};

    collegeRows.forEach(row=>{

        const major =
        row["모집단위"] ||
        "미분류";

        if(!majorMap[major]){

            majorMap[major] = [];
        }

        majorMap[major].push(row);

    });

    Object
    .entries(majorMap)
    .map(([major,rows])=>{

        const passRows =
        rows.filter(
            r=>r["최종단계"]==="합격"
        );

        const passScores =
        passRows
        .map(r=>
            Number(
                r[scoreColumn]
            )
        )
        .filter(v=>!isNaN(v));

        return {

            major,

            apply:
            rows.length,

            pass:
            passRows.length,

            rate:
            rows.length
            ?
            (
                passRows.length /
                rows.length *
                100
            )
            .toFixed(1)
            :
            "0.0",

            cutLine:
            passScores.length
            ?
            Math.max(...passScores)
            .toFixed(2)
            :
            "-"

        };

    })
    .sort(
        (a,b)=>
        Number(b.rate) -
        Number(a.rate)
    )
    .forEach(item=>{

        const tr =
        document
        .createElement("tr");

        tr.innerHTML = `

            <td>${item.major}</td>

            <td>${item.apply}</td>

            <td>${item.pass}</td>

            <td>${item.rate}%</td>

            <td>${item.cutLine}</td>

        `;

        tbody.appendChild(tr);

    });

}

function changeSimilarFilter(type){

    currentSimilarFilter = type;

    document
    .querySelectorAll(".similar-btn")
    .forEach(btn=>{

        btn.classList.remove(
            "active"
        );

    });

    if(type==="all"){

        document
        .getElementById(
            "similarAllBtn"
        )
        .classList
        .add("active");

    }
    else if(type==="pass"){

        document
        .getElementById(
            "similarPassBtn"
        )
        .classList
        .add("active");

    }
    else{

        document
        .getElementById(
            "similarFailBtn"
        )
        .classList
        .add("active");

    }

    buildSimilarStudents(
        window.currentMyScore
    );

}

function initMajorSearch(){

    const input =
    document.getElementById(
        "majorSearchInput"
    );

    input.addEventListener(
        "input",
        function(){

            const keyword =
            this.value.trim();

            if(keyword.length < 1){

                document
                .getElementById(
                    "majorSuggestions"
                )
                .style.display =
                "none";

                return;
            }

            const filteredData =
            getFilteredData();

            const majors =

            filteredData

            .filter(r=>

                r["모집단위"] &&
                r["모집단위"]
                .includes(keyword)

            )

            .map(r=>({

                college :
                r["대학명"],

                major :
                r["모집단위"]

            }))

            .slice(0,20);

            renderMajorSuggestions(
                majors
            );

        }
    );

}

function renderMajorSuggestions(
    majors
){

    const area =
    document.getElementById(
        "majorSuggestions"
    );

    if(majors.length===0){

        area.style.display =
        "none";

        return;
    }

    area.innerHTML =

    majors.map(item=>`

        <div
            class="
            major-suggestion-item
            "
            onclick="
            selectMajor(
                '${item.major}'
            )
            ">

            <div style="
                font-weight:700;
                color:#1a73e8;
            ">
                ${item.college}
            </div>

            <div style="
                font-size:14px;
                color:#5f6368;
                margin-top:3px;
            ">
                ${item.major}
            </div>

        </div>

    `).join("");

    area.style.display =
    "block";

}

function searchMajorKeyword(){

    document
.getElementById(
    "majorSuggestions"
)
.style.display =
"none";

    const keyword =

    document
    .getElementById(
        "majorSearchInput"
    )
    .value
    .trim();

    currentMajorKeyword =
            keyword;

    if(!keyword){

        alert(
            "모집단위를 입력하세요."
        );

        return;
    }

    analyzeMajorGlobal(
        keyword
    );

}

/* =========================
   모집단위 자동완성 시작
========================= */

function buildMajorSearch(){

    const input =
    document.getElementById(
        "majorSearchInput"
    );

    if(!input) return;

    input.oninput = function(){

        const keyword =
        this.value.trim();

        const area =
        document.getElementById(
            "majorSuggestions"
        );

        area.innerHTML = "";

        if(keyword.length < 1){

            area.style.display =
            "none";

            return;
        }

        const majors =

        admissionData

        .filter(r=>

            r["모집단위"] &&
            r["모집단위"]
            .includes(keyword)

        )

        .map(r=>({

            college :
            r["대학명"],

            major :
            r["모집단위"]

        }))

        .slice(0,20);

        renderMajorSuggestions(
            majors
        );

    };

}

function analyzeMajorGlobal(
    keyword
){

    const rows =

    getFilteredData()

    .filter(row=>

        row["모집단위"] &&

        row["모집단위"]
        .includes(keyword)

    );

    document
    .getElementById(
        "selectedMajorTitle"
    )
    .innerText =
    `"${keyword}" 검색 결과`;

    document
    .getElementById(
        "majorAnalysisArea"
    )
    .style.display =
    "block";

    const passRows =
    rows.filter(
        r=>r["최종단계"]==="합격"
    );

    const failRows =
    rows.filter(
        r=>r["최종단계"]!=="합격"
    );

    const collegeSet =
    new Set(
        rows.map(
            r=>r["대학명"]
        )
    );

    const yearSet =
    new Set(
        rows.map(
            r=>r["학년도"]
        )
    );

    document
    .getElementById(
        "majorTotalApply"
    )
    .innerText =
    rows.length;

    document
    .getElementById(
        "majorTotalPass"
    )
    .innerText =
    passRows.length;

    document
    .getElementById(
        "majorTotalFail"
    )
    .innerText =
    failRows.length;

    document
    .getElementById(
        "majorTotalRate"
    )
    .innerText =

    rows.length

    ?

    (
        passRows.length /
        rows.length *
        100
    ).toFixed(1)

    + "%"

    :

    "0%";

    document
    .getElementById(
        "majorCollegeCount"
    )
    .innerText =
    collegeSet.size;

    document
    .getElementById(
        "majorRecentYear"
    )
    .innerText =

    yearSet.size

    ?

    Math.max(
        ...Array.from(yearSet)
    )

    :

    "-";

    const scoreField =
    document
    .getElementById(
        "scoreFilter"
    )
    .value;

    const passScores =
    passRows
    .map(r=>
        Number(
            r[scoreField]
        )
    )
    .filter(v=>
        !isNaN(v)
    );

    const failScores =
    failRows
    .map(r=>
        Number(
            r[scoreField]
        )
    )
    .filter(v=>
        !isNaN(v)
    );

    document
    .getElementById(
        "majorAvgPassScore"
    )
    .innerText =

    passScores.length

    ?

    (
        passScores.reduce(
            (a,b)=>a+b,
            0
        )
        /
        passScores.length
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "majorAvgFailScore"
    )
    .innerText =

    failScores.length

    ?

    (
        failScores.reduce(
            (a,b)=>a+b,
            0
        )
        /
        failScores.length
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "majorMinPassScore"
    )
    .innerText =

    passScores.length

    ?

    Math.max(
        ...passScores
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "majorMaxFailScore"
    )
    .innerText =

    failScores.length

    ?

    Math.min(
        ...failScores
    ).toFixed(2)

    :

    "-";

    buildMajorCollegeTable(
        rows
    );

}

function buildMajorCollegeTable(
    rows
){

    const tbody =
    document.getElementById(
        "majorCollegeBody"
    );

    tbody.innerHTML = "";

    const majorMap = {};

    rows.forEach(row=>{

        const key =

            row["대학명"] +
            "||" +
            row["모집단위"];

        if(!majorMap[key]){

            majorMap[key] = {

                college :
                row["대학명"],

                major :
                row["모집단위"],

                apply : 0,

                pass : 0,

                fail : 0

            };
        }

        majorMap[key].apply++;

        if(
            row["최종단계"]==="합격"
        ){

            majorMap[key].pass++;

        }else{

            majorMap[key].fail++;
        }

    });

    const list =
    Object.values(
        majorMap
    );

    list.forEach(item=>{

        item.rate =

        item.apply

        ?

        item.pass /
        item.apply *
        100

        :

        0;

    });

    const sortType =

    document
    .getElementById(
        "majorSortType"
    )
    ?.value

    ||

    "college";

    switch(sortType){

        case "rate":

            list.sort(
                (a,b)=>
                b.rate-a.rate
            );

            break;

        case "apply":

            list.sort(
                (a,b)=>
                b.apply-a.apply
            );

            break;

        case "pass":

            list.sort(
                (a,b)=>
                b.pass-a.pass
            );

            break;

        default:

            list.sort((a,b)=>{

                if(
                    a.college===
                    b.college
                ){

                    return a.major
                    .localeCompare(
                        b.major,
                        "ko"
                    );

                }

                return a.college
                .localeCompare(
                    b.college,
                    "ko"
                );

            });

    }

    list.forEach(item=>{

        const tr =
        document.createElement("tr");

        tr.innerHTML = `

            <td>
                ${item.college}
            </td>

            <td>
                ${item.major}
            </td>

            <td>
                ${item.apply}
            </td>

            <td>
                ${item.pass}
            </td>

            <td>
                ${item.fail}
            </td>

            <td>
                ${item.rate.toFixed(1)}%
            </td>

        `;

        tbody.appendChild(
            tr
        );

    });

}

document
.getElementById(
    "majorSearchBtn"
)
.addEventListener(
    "click",
    searchMajorKeyword
);

document
.getElementById(
    "majorSortType"
)
.addEventListener(
    "change",
    function(){

        if(
            currentMajorKeyword
        ){

            analyzeMajorGlobal(
                currentMajorKeyword
            );

        }

    }
);

function initAdmissionFilters(){

    const typeSelect =
    document.getElementById(
        "admissionTypeFilter"
    );

    if(!typeSelect) return;

    typeSelect.innerHTML =

    `
    <option value="">
        전체 전형유형
    </option>
    `;

    const types =

    [...new Set(

        admissionData
        .map(
            r=>r["전형유형"]
        )
        .filter(v=>v)

    )]

    .sort();

    const collegeSelect =
document.getElementById(
    "admissionCollegeFilter"
);

collegeSelect.innerHTML =
`
<option value="">
전체 대학
</option>
`;

const colleges =

[...new Set(

    admissionData
    .map(r=>r["대학명"])
    .filter(v=>v)

)]

.sort();

colleges.forEach(college=>{

    const option =
    document.createElement("option");

    option.value =
    college;

    option.textContent =
    college;

    collegeSelect.appendChild(
        option
    );

});

    types.forEach(type=>{

        const option =
        document.createElement(
            "option"
        );

        option.value =
        type;

        option.textContent =
        type;

        typeSelect.appendChild(
            option
        );

    });

    typeSelect.onchange =
    function(){

    updateDetailTypeFilter();

    analyzeAdmission();

    };

    document
    .getElementById(
        "detailTypeFilter"
    )
    .onchange =
    analyzeAdmission;

    document
    .getElementById(
        "admissionCollegeFilter"
    )
    .onchange =
    analyzeAdmission;

    document
    .getElementById(
        "admissionSort"
    )
    .onchange =
    analyzeAdmission;

}

function updateDetailTypeFilter(){

    const type =

    document
    .getElementById(
        "admissionTypeFilter"
    )
    .value;

    const detailSelect =

    document
    .getElementById(
        "detailTypeFilter"
    );

    detailSelect.innerHTML =

    `
    <option value="">
        전체 세부유형
    </option>
    `;

    let rows =
    admissionData;

    if(type){

        rows = rows.filter(
            r=>
            r["전형유형"]===type
        );

    }

    const details =

    [...new Set(

        rows
        .map(
            r=>r["세부유형"]
        )
        .filter(v=>v)

    )]

    .sort();

    details.forEach(detail=>{

        const option =
        document.createElement(
            "option"
        );

        option.value =
        detail;

        option.textContent =
        detail;

        detailSelect.appendChild(
            option
        );
  
    });

analyzeAdmission();

}

function analyzeAdmission(){

    const type =

    document
    .getElementById(
        "admissionTypeFilter"
    )
    .value;

    const detail =

    document
    .getElementById(
        "detailTypeFilter"
    )
    .value;

    const detailKeyword =

    document
    .getElementById(
        "detailKeyword"
    )
    .value
    .trim();

    const college =

    document
    .getElementById(
        "admissionCollegeFilter"
    )
.value;

    let rows =
    getFilteredData();

    if(type){

        rows = rows.filter(
            r=>
            r["전형유형"]===type
        );

    }

    if(detail){

        rows = rows.filter(
            r=>
            r["세부유형"]===detail
        );

    }

    if(detailKeyword){

        rows = rows.filter(
            r=>

            r["세부유형"]

            &&

            r["세부유형"]
            .includes(
                detailKeyword
            )

        );

    }

    if(college){

        rows = rows.filter(
            r=>
        r["대학명"]===college
        );

    }

    currentAdmissionRows =
    rows;

    document
    .getElementById(
        "admissionAnalysisArea"
    )
    .style.display =
    "block";

    let title = "";

    if(type){

        title += type;

    }

    if(detail){

        title +=
        " > " +
        detail;

    }

    if(detailKeyword){

        title +=
        " > " +
        detailKeyword;

    }

    if(college){

        title +=
        " > " +
        college;

    }

    document
    .getElementById(
        "selectedAdmissionTitle"
    )
    .innerText =
    title || "전체 전형";

    updateAdmissionCards(
        rows
    );

    buildAdmissionTable(
        rows
    );

}

function buildAdmissionTable(
    rows
){

    const tbody =
    document.getElementById(
        "admissionTableBody"
    );

    tbody.innerHTML = "";

    const map = {};

    rows.forEach(row=>{

        const key =

                row["대학명"]
                + "||" +
                row["모집단위"]
                + "||" +
                row["세부유형"];

        if(!map[key]){

            map[key] = {

                college :
                row["대학명"],

                major :
                row["모집단위"],

                detail :
                row["세부유형"],

                apply : 0,

                pass : 0,

                fail : 0

            };

        }

        map[key].apply++;

        if(
            row["최종단계"]==="합격"
        ){

            map[key].pass++;

        }else{

            map[key].fail++;

        }

    });

    const list =
    Object.values(map);

    const sortType =

    document
    .getElementById(
        "admissionSort"
    )
    .value;

    if(sortType==="apply"){

    list.sort(
        (a,b)=>

        b.apply -
        a.apply
    );

}

else if(
    sortType==="pass"
){

    list.sort(
        (a,b)=>

        b.pass -
        a.pass
    );

}

else if(
    sortType==="rate"
){

    list.sort(
        (a,b)=>

        (b.pass/b.apply)
        -
        (a.pass/a.apply)
    );

}

else{

    list.sort((a,b)=>{

        if(
            a.college===
            b.college
        ){

            return a.major
            .localeCompare(
                b.major,
                "ko"
            );

        }

        return a.college
        .localeCompare(
            b.college,
            "ko"
        );

    });

}

    list.forEach(item=>{

        const rate =

        item.apply

        ?

        (
            item.pass /
            item.apply *
            100
        ).toFixed(1)

        :

        "0.0";

        const tr =
        document.createElement(
            "tr"
        );

        tr.innerHTML = `

           <td>${item.college}</td>

            <td>${item.major}</td>

            <td>${item.detail}</td>

            <td>${item.apply}</td>

            <td>${item.pass}</td>

            <td>${item.fail}</td>

            <td>${rate}%</td>
        `;

        tbody.appendChild(
            tr
        );

    });

}

function updateAdmissionCards(
    rows
){

    const scoreField =

    document
    .getElementById(
        "scoreFilter"
    )
    .value;

    const passRows =
    rows.filter(
        r=>r["최종단계"]==="합격"
    );

    const failRows =
    rows.filter(
        r=>r["최종단계"]!=="합격"
    );

    const passScores =
    passRows
    .map(r=>
        Number(
            r[scoreField]
        )
    )
    .filter(v=>!isNaN(v));

    const failScores =
    failRows
    .map(r=>
        Number(
            r[scoreField]
        )
    )
    .filter(v=>!isNaN(v));

    const collegeSet =
    new Set(
        rows.map(
            r=>r["대학명"]
        )
    );

    const majorSet =
    new Set(
        rows.map(
            r=>r["모집단위"]
        )
    );

    document
    .getElementById(
        "admissionApplyCount"
    )
    .innerText =
    rows.length;

    document
    .getElementById(
        "admissionPassCount"
    )
    .innerText =
    passRows.length;

    document
    .getElementById(
        "admissionFailCount"
    )
    .innerText =
    failRows.length;

    document
    .getElementById(
        "admissionPassRate"
    )
    .innerText =

    rows.length

    ?

    (
        passRows.length /
        rows.length *
        100
    ).toFixed(1)

    + "%"

    :

    "0%";

    document
    .getElementById(
        "admissionCollegeCount"
    )
    .innerText =
    collegeSet.size;

    document
    .getElementById(
        "admissionMajorCount"
    )
    .innerText =
    majorSet.size;

    document
    .getElementById(
        "admissionAvgPass"
    )
    .innerText =

    passScores.length

    ?

    (
        passScores.reduce(
            (a,b)=>a+b,
            0
        )
        /
        passScores.length
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "admissionAvgFail"
    )
    .innerText =

    failScores.length

    ?

    (
        failScores.reduce(
            (a,b)=>a+b,
            0
        )
        /
        failScores.length
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "admissionMinPass"
    )
    .innerText =

    passScores.length

    ?

    Math.max(
        ...passScores
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "admissionMaxFail"
    )
    .innerText =

    failScores.length

    ?

    Math.min(
        ...failScores
    ).toFixed(2)

    :

    "-";

}

function initRegionAnalysis(){

    const select =
    document.getElementById(
        "regionAnalysisFilter"
    );

    if(!select) return;

    select.innerHTML =
    `
    <option value="">
        전체 지역
    </option>
    `;

    const regions =

    [...new Set(

        admissionData
        .map(r=>r["지역"])
        .filter(v=>v)

    )]

    .sort();

    regions.forEach(region=>{

        const option =
        document.createElement(
            "option"
        );

        option.value =
        region;

        option.textContent =
        region;

        select.appendChild(
            option
        );

    });

    select.onchange =
    analyzeRegion;
    analyzeRegion();
}

function analyzeRegion(){

    const region =

    document
    .getElementById(
        "regionAnalysisFilter"
    )
    .value;

    let rows =
    getFilteredData();

    if(region){

        rows = rows.filter(
            r=>
            r["지역"]===region
        );

    }

    document
    .getElementById(
        "selectedRegionTitle"
    )
    .innerText =

    region ||

    "전체 지역";

    const passRows =
    rows.filter(
        r=>r["최종단계"]==="합격"
    );

    const failRows =
    rows.filter(
        r=>r["최종단계"]!=="합격"
    );

    document
    .getElementById(
        "regionApplyCount"
    )
    .innerText =
    rows.length;

    document
    .getElementById(
        "regionPassCount"
    )
    .innerText =
    passRows.length;

    document
    .getElementById(
        "regionFailCount"
    )
    .innerText =
    failRows.length;

    document
    .getElementById(
        "regionPassRate"
    )
    .innerText =

    rows.length

    ?

    (
        passRows.length /
        rows.length *
        100
    ).toFixed(1)

    + "%"

    :

    "0%";

    const scoreField =
    document
    .getElementById("scoreFilter")
    .value;

    const passScores =
    passRows
    .map(r=>Number(r[scoreField]))
    .filter(v=>!isNaN(v));

    const failScores =
    failRows
    .map(r=>Number(r[scoreField]))
    .filter(v=>!isNaN(v));

    document
    .getElementById(
        "regionAvgPass"
    )
    .innerText =

    passScores.length

    ?

    (
        passScores
        .reduce(
            (a,b)=>a+b,
            0
        )
        /
        passScores.length
    ).toFixed(2)

    :

    "-";

    document
    .getElementById(
        "regionAvgFail"
    )
    .innerText =

    failScores.length

    ?

    (
        failScores
        .reduce(
            (a,b)=>a+b,
            0
        )
        /
        failScores.length
    ).toFixed(2)

    :

    "-";

    buildRegionCollegeTable(
        rows
    );

}

function buildRegionCollegeTable(
    rows
){

    const tbody =
    document.getElementById(
        "regionCollegeBody"
    );

    tbody.innerHTML = "";

    const map = {};

    rows.forEach(row=>{

        const college =
        row["대학명"];

        if(!map[college]){

            map[college] = {

                college,

                apply : 0,

                pass : 0,

                fail : 0

            };

        }

        map[college].apply++;

        if(
            row["최종단계"]==="합격"
        ){

            map[college].pass++;

        }else{

            map[college].fail++;

        }

    });

    const list =
    Object.values(map);

    list.sort(
        (a,b)=>
        b.apply-a.apply
    );

    list.forEach(item=>{

        const rate =

        item.apply

        ?

        (
            item.pass /
            item.apply *
            100
        ).toFixed(1)

        :

        "0.0";

        const tr =
        document.createElement(
            "tr"
        );

        tr.innerHTML = `

            <td>
                ${item.college}
            </td>

            <td>
                ${item.apply}
            </td>

            <td>
                ${item.pass}
            </td>

            <td>
                ${item.fail}
            </td>

            <td>
                ${rate}%
            </td>

        `;

        tbody.appendChild(
            tr
        );

    });

}

