//=============================
// 0. 초기 세팅
//=============================
//함수 : DOM 선택자
function $(parameter, startNode) {
  if (!startNode) return document.querySelector(parameter)
  else return startNode.querySelector(parameter)
}
function $$(parameter, startNode) {
  if (!startNode) return document.querySelectorAll(parameter)
  else return startNode.querySelectorAll(parameter)
}
//함수 : 일시정지 (출처: https://inpa.tistory.com/entry/JS-📚-자바스크립트에-sleep-wait-대기-함수-쓰기)
async function wait(sec) {
  let start = Date.now(), now = start;
  while (now - start < sec * 1000) {
      now = Date.now();
  }
}
//변수 : 데이터 불러오기 간략화
let PROP = dataObj.wsProperty
let CONT = dataObj.wsContents
const dataObjDefault = structuredClone(dataObj)//사본
//요일 배열(요일 구하는 용)
const WEEKARR = ['일','월','화','수','목','금','토']
//주제 배열
const SUBJECTARR = ["총류","철학","종교","사회과학","자연과학","기술과학","예술","언어","문학","역사"]
//합계 배열(일계, 월계, 연계)
const TOTALARR = ['일계','월계','연계']
//기타 배열(자료 입력용)
const 비도서typeARR = [["category","소제목"],["item","부록명"]]
const subTotalsRefARR = [["도서","도서 합계"],["비도서","비도서 합계"],["all","도서+비도서"],["none","자유 입력"]]

//색상 선택 태그 설정
Coloris({
  theme:'polaroid',
  alpha:false,
  closeButton:true,
  closeLabel:"완료",
  swatches:[
    '#0000FF',//파랑
    '#FF0000',//빨강
    '#A020F0',//보라
    '#006600',//초록
    '#FF00FF',//분홍
    '#F58700'//주황
  ]
})


//=============================
// 1. 입력내용 입출력 및 임시저장 관련
//=============================
//1-1. 입력
let setContent = (input) => {
  //별도 입력값을 지정하지 않으면 디폴트값 불러오기
  if (!input) input = structuredClone(dataObjDefault)
//PROP
  //1. 파일 기본 설정
  $("#input_PROP_library").value = input.wsProperty.library.replaceAll("\n","\\n")
  if (input.wsProperty.date[0] !== "" && input.wsProperty.date[1] !== "") {
      $("#input_PROP_date").value = input.wsProperty.date[0] + "-" + ("0" + input.wsProperty.date[1]).slice(-2)
  } else {
    let thisTime = new Date()
    $("#input_PROP_date").value = thisTime.getFullYear().toString() + "-" + ("0" + (thisTime.getMonth()+1).toString()).slice(-2)
  }
  $("#input_PROP_fontName").value = input.wsProperty.fontName
    $("#input_PROP_fontSize_title").value = input.wsProperty.fontSize.title.toString()
    $("#input_PROP_fontSize_large").value = input.wsProperty.fontSize.large.toString()
    $("#input_PROP_fontSize_semiLarge").value = input.wsProperty.fontSize.semiLarge.toString()
    $("#input_PROP_fontSize_middle").value = input.wsProperty.fontSize.middle.toString()
    $("#input_PROP_fontSize_small").value = input.wsProperty.fontSize.small.toString()
  $("#input_PROP_rowHeight_first").value = input.wsProperty.rowHeight.first.toString()
  $("#input_PROP_rowHeight_other").value = input.wsProperty.rowHeight.other.toString()
  $("#input_PROP_colWidth").value = input.wsProperty.colWidth.toString()
  $("#input_PROP_zoom").value = input.wsProperty.zoom.toString()
  //2. 요일별 설정
  for (let dateChar in input.wsProperty.dateType) {
    let dateObj = input.wsProperty.dateType[dateChar]
    //마지막 주 설정 여부
    $("#input_PROP_dateType_" + dateChar + "_useLast").checked = dateObj.useLast
      $("#input_PROP_dateType_" + dateChar + "_useLast").dispatchEvent(new Event("change"))
    for (let objKey in dateObj) {
      if (typeof dateObj[objKey] === "object") {
        //문구
        $("#input_PROP_dateType_" + dateChar + "_" + objKey + "_desc").value = dateObj[objKey].desc.replaceAll("\n","\\n")
        //색상
        let colorText = (dateObj[objKey].color !== "") ? dateObj[objKey].color : "0000FF"
        $("#input_PROP_dateType_" + dateChar + "_" + objKey + "_color").value = "#" + colorText
        $("#input_PROP_dateType_" + dateChar + "_" + objKey + "_color").dispatchEvent(new Event('input', { bubbles: true }));
      }
    }
  }
//CONT
  //1. 관외대출
  $("#input_CONT_관외대출_title").value = input.wsContents.관외대출.title.replaceAll("\n","\\n")
  setElement("관외대출",input)

  //2. 이용자
  $("#input_CONT_이용자_title").value = input.wsContents.이용자.title.replaceAll("\n","\\n")
  setElement("이용자_col",input)
  setElement("이용자_row",input)
  if ($("#input_CONT_이용자_wide_" + input.wsContents.이용자.wide))
    $("#input_CONT_이용자_wide_" + input.wsContents.이용자.wide).checked = true

  //3. 야간
  $("#input_CONT_야간_enabled").checked = (input.wsContents.야간.enabled === 1) ? true : false
    $("#input_CONT_야간_enabled").dispatchEvent(new Event("change"))
  $("#input_CONT_야간_title").value = input.wsContents.야간.title.replaceAll("\n","\\n")
  $("#input_CONT_야간_study").checked = (input.wsContents.야간.study === 1) ? true : false
    $("#input_CONT_야간_study").dispatchEvent(new Event("change"))
  $("#input_CONT_야간_studyName").value = input.wsContents.야간.studyName.replaceAll("\n","\\n")

  //4. 특성화
  $("#input_CONT_특성화_enabled").checked = (input.wsContents.특성화.enabled === 1) ? true : false
    $("#input_CONT_특성화_enabled").dispatchEvent(new Event("change"))
  $("#input_CONT_특성화_title").value = input.wsContents.특성화.title.replaceAll("\n","\\n")
  $("#input_CONT_특성화_name").value = input.wsContents.특성화.name.replaceAll("\n","\\n")

  //5. 소장자료
  $("#input_CONT_소장자료_title").value = input.wsContents.소장자료.title.replaceAll("\n","\\n")
  setElement("소장자료_자료실",input)
  setElement("소장자료_비도서",input)
  $("#input_CONT_소장자료_totalName").value = input.wsContents.소장자료.totalName.replaceAll("\n","\\n")
  setElement("소장자료_subTotal",input)

  //6. 회원가입
  $("#input_CONT_회원가입_title").value = input.wsContents.회원가입.title.replaceAll("\n","\\n")
  setElement("회원가입",input)

  //7. 강좌
  $("#input_CONT_강좌_title").value = input.wsContents.강좌.title.replaceAll("\n","\\n")
  setElement("강좌",input)

  //8. 기타
  $("#input_CONT_기타_title").value = input.wsContents.기타.title.replaceAll("\n","\\n")
  $("#input_CONT_기타_lines").value = input.wsContents.기타.lines.toString()
}


//1-2. 출력
let getContent = () => {
  let output = {wsProperty:{},wsContents:{}}
//PROP
  //1. 파일 기본 설정
  output.wsProperty.library = $("#input_PROP_library").value.replaceAll("\\n","\n")
  output.wsProperty.date = []
    output.wsProperty.date[0] = $("#input_PROP_date").value.split("-")[0]
    output.wsProperty.date[1] = Number($("#input_PROP_date").value.split("-")[1]).toString()
  output.wsProperty.fontName = $("#input_PROP_fontName").value
  output.wsProperty.fontSize = {}
    output.wsProperty.fontSize.title = Number($("#input_PROP_fontSize_title").value)
    output.wsProperty.fontSize.large = Number($("#input_PROP_fontSize_large").value)
    output.wsProperty.fontSize.semiLarge = Number($("#input_PROP_fontSize_semiLarge").value)
    output.wsProperty.fontSize.middle = Number($("#input_PROP_fontSize_middle").value)
    output.wsProperty.fontSize.small = Number($("#input_PROP_fontSize_small").value)
  output.wsProperty.rowHeight = {}
    output.wsProperty.rowHeight.first = Number($("#input_PROP_rowHeight_first").value)
    output.wsProperty.rowHeight.other = Number($("#input_PROP_rowHeight_other").value)
  output.wsProperty.colWidth = Number($("#input_PROP_colWidth").value)
  output.wsProperty.zoom = Number($("#input_PROP_zoom").value)
  //2. 요일별 설정
  output.wsProperty.dateType = {}
  WEEKARR.forEach(dateChar => {
    output.wsProperty.dateType[dateChar] = {}
    let dateObj = output.wsProperty.dateType[dateChar]
    dateObj.useLast = $("#input_PROP_dateType_" + dateChar + "_useLast").checked
    dateObj.ordinary = {}
    dateObj.ordinary.desc = $("#input_PROP_dateType_" + dateChar + "_ordinary_desc").value.replaceAll("\\n","\n")
    if (dateObj.ordinary.desc !== "") {
      dateObj.ordinary.color = $("#input_PROP_dateType_" + dateChar + "_ordinary_color").value.replace("#","")
    } else dateObj.ordinary.color = ""
    dateObj.last = {}
    dateObj.last.desc = $("#input_PROP_dateType_" + dateChar + "_last_desc").value.replaceAll("\\n","\n")
    if (dateObj.last.desc !== "") {
      dateObj.last.color = $("#input_PROP_dateType_" + dateChar + "_last_color").value.replace("#","")
    } else dateObj.last.color = ""
  })
//CONT
  //1. 관외대출
  output.wsContents.관외대출 = {}
  output.wsContents.관외대출.title = $("#input_CONT_관외대출_title").value.replaceAll("\\n","\n")
  output.wsContents.관외대출.cols = []
  $$(".input_CONT_관외대출_col").forEach(col => {
    output.wsContents.관외대출.cols.push(col.value)
  })
  //2. 이용자
  output.wsContents.이용자 = {}
  output.wsContents.이용자.title = $("#input_CONT_이용자_title").value.replaceAll("\\n","\n")
  output.wsContents.이용자.cols = []
  $$(".input_CONT_이용자_col").forEach(col => {
    output.wsContents.이용자.cols.push(col.value)
  })
  output.wsContents.이용자.rows = []
  $$(".input_CONT_이용자_row").forEach(row => {
    output.wsContents.이용자.rows.push(row.value)
  })
  output.wsContents.이용자.wide = $("input[name='input_CONT_이용자_wide']:checked").value
  //3. 야간
  output.wsContents.야간 = {}
  output.wsContents.야간.title = $("#input_CONT_야간_title").value.replaceAll("\\n","\n")
  output.wsContents.야간.enabled = ($("#input_CONT_야간_enabled").checked === true) ? 1 : 0
  output.wsContents.야간.study = ($("#input_CONT_야간_study").checked === true) ? 1 : 0
  output.wsContents.야간.studyName = $("#input_CONT_야간_studyName").value.replaceAll("\\n","\n")
  //4. 특성화
  output.wsContents.특성화 = {}
  output.wsContents.특성화.title = $("#input_CONT_특성화_title").value.replaceAll("\\n","\n")
  output.wsContents.특성화.name = $("#input_CONT_특성화_name").value.replaceAll("\\n","\n")
  output.wsContents.특성화.enabled = ($("#input_CONT_특성화_enabled").checked === true) ? 1 : 0
  //4-2. 이용자 분류 초과 체크
  let tempRow = output.wsContents.이용자.rows.length
  let tempEnabled = [output.wsContents.야간.enabled,output.wsContents.특성화.enabled]
  let wide = output.wsContents.이용자.wide
  let tooMuch분류 = 0
  if (tempEnabled[0] === 0 && tempEnabled[1] === 0) {//야간 X, 특성화 X (빈칸 16)
    if (wide === "11" && tempRow > 14) tooMuch분류 = 1
    else if (wide === "21" && tempRow > 7) tooMuch분류 = 1
    else if (wide === "22" && tempRow > 6) tooMuch분류 = 1
    else if (wide === "32" && tempRow > 4) tooMuch분류 = 1
  } else if (tempEnabled[0] === 0 && tempEnabled[1] === 1) {//야간 X, 특성화 O (빈칸 14)
    if (wide === "11" && tempRow > 12) tooMuch분류 = 1
    else if (wide === "21" && tempRow > 6) tooMuch분류 = 1
    else if (wide === "22" && tempRow > 5) tooMuch분류 = 1
    else if (wide === "32" && tempRow > 3) tooMuch분류 = 1
  } else if (tempEnabled[0] === 1 && tempEnabled[1] === 0) {//야간 O, 특성화 X (빈칸 12)
    if (wide === "11" && tempRow > 10) tooMuch분류 = 1
    else if (wide === "21" && tempRow > 5) tooMuch분류 = 1
    else if (wide === "22" && tempRow > 4) tooMuch분류 = 1
    else if (wide === "32" && tempRow > 2) tooMuch분류 = 1
  } else if (tempEnabled[0] === 1 && tempEnabled[1] === 1) {//야간 O, 특성화 O (빈칸 10)
    if (wide === "11" && tempRow > 8) tooMuch분류 = 1
    else if (wide === "21" && tempRow > 4) tooMuch분류 = 1
    else if (wide === "22" && tempRow > 3) tooMuch분류 = 1
    else if (wide === "32" && tempRow > 2) tooMuch분류 = 1
  }
  if (tooMuch분류 === 1) {
    alert("* 오류 : 이용 분류가 너무 많습니다. 안내문을 참고하여 다시 설정해주세요.")
    throw new Error('입력값 수집 정지됨 : 이용자 분류 수량 초과')
  }
  //5. 소장자료
  output.wsContents.소장자료 = {}
  output.wsContents.소장자료.title = $("#input_CONT_소장자료_title").value.replaceAll("\\n","\n")
  output.wsContents.소장자료.자료실 = []
  $$(".input_CONT_소장자료_자료실").forEach(col => {
    output.wsContents.소장자료.자료실.push(col.value)
  })
  output.wsContents.소장자료.비도서 = []
  output.wsContents.소장자료.비도서type = []
  $$(".input_CONT_소장자료_비도서").forEach((col,i) => {
    output.wsContents.소장자료.비도서.push(col.value)
    output.wsContents.소장자료.비도서type.push($("#input_CONT_소장자료_비도서type_" + (i+1).toString()).value)
  })
  output.wsContents.소장자료.totalName = $("#input_CONT_소장자료_totalName").value.replaceAll("\\n","\n")
  output.wsContents.소장자료.subTotals = []
  output.wsContents.소장자료.subTotalsRef = []
  $$(".input_CONT_소장자료_subTotal").forEach((col,i) => {
    output.wsContents.소장자료.subTotals.push(col.value)
    output.wsContents.소장자료.subTotalsRef.push($("#input_CONT_소장자료_subTotalsRef_" + (i+1).toString()).value)
  })
  //6. 회원가입
  output.wsContents.회원가입 = {}
  output.wsContents.회원가입.title = $("#input_CONT_회원가입_title").value.replaceAll("\\n","\n")
  output.wsContents.회원가입.groups = []
  $$(".input_CONT_회원가입_group").forEach((group,i) => {
    let groupObj = {}
    groupObj.name = $("#input_CONT_회원가입_group_name_" + (i+1).toString()).value
    groupObj.items = $("#input_CONT_회원가입_group_items_" + (i+1).toString()).value.split(",").map(item => item.trim())
    groupObj.allTotal = $("#input_CONT_회원가입_group_allTotal_" + (i+1).toString()).checked
    output.wsContents.회원가입.groups.push(groupObj)
  })
  //7. 강좌
  output.wsContents.강좌 = {}
  output.wsContents.강좌.title = $("#input_CONT_강좌_title").value.replaceAll("\\n","\n")
  output.wsContents.강좌.groups = []
  output.wsContents.강좌.lines = 0
  $$(".input_CONT_강좌_group").forEach((group,i) => {
    let groupObj = {}
    groupObj.name = $("#input_CONT_강좌_group_name_" + (i+1).toString()).value.replaceAll("\\n","\n")
    groupObj.lines = 0
    if ($("#input_CONT_강좌_group_sub_" + (i+1).toString()).checked) {
      groupObj.subGroup = []
      let subNameArr = $("#input_CONT_강좌_group_subname_" + (i+1).toString()).value.split(",").map(item => item.trim())
      let subLinesArr = $("#input_CONT_강좌_group_lines_" + (i+1).toString()).value.split(",").map(item => item.trim())
      //쉼표 나누기가 제대로 되지 않으면 경고 후 종료
      if (subNameArr.length !== subLinesArr.length) {
        alert("* 오류 : 강좌 및 행사 현황 - " + (i+1).toString() + "번 범주의 추가범주/입력칸수의 쉼표 수가 일치하지 않습니다.")
        throw new Error('입력값 수집 정지됨 : 입력값 오류')
      } else {
        subNameArr.forEach((x,j) => {
          let subGroupObj = {}
          subGroupObj.name = subNameArr[j].replaceAll("\\n","\n")
          subGroupObj.lines = Number(subLinesArr[j])
          groupObj.subGroup.push(subGroupObj)

          groupObj.lines += subGroupObj.lines
          output.wsContents.강좌.lines += subGroupObj.lines
        })
      }
    } else {
      //쉼표가 있으면 경고 후 종료
      if ($("#input_CONT_강좌_group_lines_" + (i+1).toString()).value.indexOf(",") >= 0) {
        alert("* 오류 : 강좌 및 행사 현황 : " + (i+1).toString() + "번 범주 입력칸수에서 쉼표를 제거해주세요.")
        throw new Error('입력값 수집 정지됨 : 입력값 오류')
      } else {
        groupObj.lines = Number($("#input_CONT_강좌_group_lines_" + (i+1).toString()).value)
        output.wsContents.강좌.lines += groupObj.lines
      }
    }
    output.wsContents.강좌.groups.push(groupObj)
  })
  //8. 기타
  output.wsContents.기타 = {}
  output.wsContents.기타.title = $("#input_CONT_기타_title").value.replaceAll("\\n","\n")
  output.wsContents.기타.lines = Number($("#input_CONT_기타_lines").value)

  return output
}

//1-3. 임시저장
let saveContent = async () => {
  //현재 입력값 불러오기
  let forSave = getContent()
  //입력값 저장하기
  localStorage.setItem("savedData", JSON.stringify(forSave))
  //저장 아이콘 표시
  $("#saveAlarm").classList.add("animated")
  $("#saveAlarm").onanimationend = () => {
    $("#saveAlarm").classList.remove("animated")
  }
}

//1-4. 임시저장 불러오기
let loadContent = () => {
  let output = localStorage.getItem("savedData")
  if (output === null) {
    return null
  } else {
    return JSON.parse(output)
  }
}

//=============================
// 2. 입력란 설정
//=============================
let setElement = (targetEl,cmd) => {
  let append//DOM 추가용 함수

  switch (targetEl) {
    case "관외대출":
      //추가 함수 준비
      append = (num,val="") => {
        let newEl = document.createElement('p')
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "자료실" + (num+1).toString() + ": "
          newEl.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_관외대출_col_" + (num+1).toString()
          newEl3.classList.add("input_CONT_관외대출_col","w3-border","long")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl.appendChild(newEl3)
        return newEl
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_관외대출").childElementCount
          $("#div_관외대출").appendChild(append(num))
          break
        case "remove":
          if ($("#div_관외대출").childElementCount > 1)//관외대출 자료실 : 최소 1개 이상
            $("#div_관외대출").removeChild($("#div_관외대출").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_관외대출").firstElementChild)
              $("#div_관외대출").removeChild($("#div_관외대출").lastElementChild)
            cmd.wsContents.관외대출.cols.forEach((name,i) => {
              $("#div_관외대출").appendChild(append(i,name.replaceAll("\n","\\n")))
            })
          }
          break
      }
      break
    
    case "이용자_col":
      //추가 함수 준비
      append = (num,val="") => {
        let newEl = document.createElement('p')
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "이용자" + (num+1).toString() + ": "
          newEl.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_이용자_col_" + (num+1).toString()
          newEl3.classList.add("input_CONT_이용자_col","w3-border","long")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl.appendChild(newEl3)
        return newEl
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_이용자_col").childElementCount
          $("#div_이용자_col").appendChild(append(num))
          break
        case "remove":
          if ($("#div_이용자_col").childElementCount > 1)//이용자 구분 : 최소 1개 이상
            $("#div_이용자_col").removeChild($("#div_이용자_col").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_이용자_col").firstElementChild)
              $("#div_이용자_col").removeChild($("#div_이용자_col").lastElementChild)
            cmd.wsContents.이용자.cols.forEach((name,i) => {
              $("#div_이용자_col").appendChild(append(i,name.replaceAll("\n","\\n")))
            })
          }
          break
      }
      break
    
    case "이용자_row":
      //추가 함수 준비
      append = (num,val="") => {
        let newEl = document.createElement('p')
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "분류" + (num+1).toString() + ": "
          newEl.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_이용자_row_" + (num+1).toString()
          newEl3.classList.add("input_CONT_이용자_row","w3-border","long")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl.appendChild(newEl3)
        return newEl
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_이용자_row").childElementCount
          $("#div_이용자_row").appendChild(append(num))
          break
        case "remove":
          if ($("#div_이용자_row").childElementCount > 1)//이용 분류 : 최소 1개 이상
            $("#div_이용자_row").removeChild($("#div_이용자_row").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_이용자_row").firstElementChild)
              $("#div_이용자_row").removeChild($("#div_이용자_row").lastElementChild)
            cmd.wsContents.이용자.rows.forEach((name,i) => {
              $("#div_이용자_row").appendChild(append(i,name.replaceAll("\n","\\n")))
            })
          }
          break
      }
      break
    
    case "소장자료_자료실":
      //추가 함수 준비
      append = (num,val="") => {
        let newEl = document.createElement('p')
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "자료실" + (num+1).toString() + ": "
          newEl.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_소장자료_자료실_" + (num+1).toString()
          newEl3.classList.add("input_CONT_소장자료_자료실","w3-border","long")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl.appendChild(newEl3)
        return newEl
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_소장자료_자료실").childElementCount
          $("#div_소장자료_자료실").appendChild(append(num))
          break
        case "remove":
          if ($("#div_소장자료_자료실").childElementCount > 1)//자료실 : 최소 1개 이상
            $("#div_소장자료_자료실").removeChild($("#div_소장자료_자료실").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_소장자료_자료실").firstElementChild)
              $("#div_소장자료_자료실").removeChild($("#div_소장자료_자료실").lastElementChild)
            cmd.wsContents.소장자료.자료실.forEach((name,i) => {
              $("#div_소장자료_자료실").appendChild(append(i,name.replaceAll("\n","\\n")))
            })
          }
          break
      }
      break
    
    case "소장자료_비도서":
      //추가 함수 준비
      append = (num,val="",val2="") => {
        let newEl = document.createElement('p')
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "항목" + (num+1).toString() + ": "
          newEl.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_소장자료_비도서_" + (num+1).toString()
          newEl3.classList.add("input_CONT_소장자료_비도서","w3-border","middle")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl.appendChild(newEl3)
        let newEl4 = document.createElement('select')
          newEl4.id = "input_CONT_소장자료_비도서type_" + (num+1).toString()
            let option
            option = document.createElement('option')
            option.disabled = true
            if (option.value === "") option.selected = true
            option.innerHTML = "항목 구분"
            newEl4.append(option)
            비도서typeARR.forEach(arr => {
              option = document.createElement('option')
              option.value = arr[0]
              option.innerHTML = arr[1]
              if (option.value === val2) option.selected = true
              newEl4.append(option)
            })
          newEl.appendChild(newEl4)
        return newEl
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_소장자료_비도서").childElementCount
          $("#div_소장자료_비도서").appendChild(append(num))
          break
        case "remove":
          if ($("#div_소장자료_비도서").childElementCount > 0)//비도서는 없을 수도 있음
            $("#div_소장자료_비도서").removeChild($("#div_소장자료_비도서").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_소장자료_비도서").firstElementChild)
              $("#div_소장자료_비도서").removeChild($("#div_소장자료_비도서").lastElementChild)
            cmd.wsContents.소장자료.비도서.forEach((name,i) => {
              $("#div_소장자료_비도서").appendChild(append(i,name.replaceAll("\n","\\n"),cmd.wsContents.소장자료.비도서type[i]))
            })
          }
          break
      }
      break
    
    case "소장자료_subTotal":
      //추가 함수 준비
      append = (num,val="",val2="") => {
        let newEl = document.createElement('p')
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "소계" + (num+1).toString() + ": "
          newEl.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_소장자료_subTotal_" + (num+1).toString()
          newEl3.classList.add("input_CONT_소장자료_subTotal","w3-border","middle")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl.appendChild(newEl3)
        let newEl4 = document.createElement('select')
          newEl4.id = "input_CONT_소장자료_subTotalsRef_" + (num+1).toString()
            let option
            option = document.createElement('option')
            option.disabled = true
            if (option.value === "") option.selected = true
            option.innerHTML = "참고 항목"
            newEl4.append(option)
            subTotalsRefARR.forEach(arr => {
              option = document.createElement('option')
              option.value = arr[0]
              option.innerHTML = arr[1]
              if (option.value === val2) option.selected = true
              newEl4.append(option)
            })
          newEl.appendChild(newEl4)
        return newEl
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_소장자료_subTotal").childElementCount
          $("#div_소장자료_subTotal").appendChild(append(num))
          break
        case "remove":
          if ($("#div_소장자료_subTotal").childElementCount > 0)//별도 소계는 없을 수도 있음
            $("#div_소장자료_subTotal").removeChild($("#div_소장자료_subTotal").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_소장자료_subTotal").firstElementChild)
              $("#div_소장자료_subTotal").removeChild($("#div_소장자료_subTotal").lastElementChild)
            cmd.wsContents.소장자료.subTotals.forEach((name,i) => {
              $("#div_소장자료_subTotal").appendChild(
                append(i,
                  name.replaceAll("\n","\\n"),
                  cmd.wsContents.소장자료.subTotalsRef[i])
              )
            })
          }
          break
      }
      break
    
    case "회원가입":
      //추가 함수 준비
      append = (num,val="",val2="",val3=false) => {
        let output = document.createElement('div')
          output.classList.add("input_CONT_회원가입_group")
        //1번줄
        let newEl1 = document.createElement('p')
          output.appendChild(newEl1)
        let newEl2 = document.createElement('span')
          newEl2.classList.add("inputHeader")
          newEl2.innerHTML = "범주" + (num+1).toString() + ": "
          newEl1.appendChild(newEl2)
        let newEl3 = document.createElement('input')
          newEl3.id = "input_CONT_회원가입_group_name_" + (num+1).toString()
          newEl3.classList.add("input_CONT_회원가입_group_name","w3-border","long")
          newEl3.type = "text"
          newEl3.value = val
          newEl3.required = true
          newEl1.appendChild(newEl3)
        //2번줄
        let newEl4 = document.createElement('p')
          output.appendChild(newEl4)
        let newEl5 = document.createElement('span')
          newEl5.classList.add("inputHeader")
          newEl5.innerHTML = "세부구분: "
          newEl4.appendChild(newEl5)
        let newEl6 = document.createElement('input')
          newEl6.id = "input_CONT_회원가입_group_items_" + (num+1).toString()
          newEl6.classList.add("input_CONT_회원가입_group_items","w3-border","long")
          newEl6.type = "text"
          newEl6.value = val2
          newEl6.required = true
          newEl4.appendChild(newEl6)
        //3번줄
        let newEl7 = document.createElement('p')
          output.appendChild(newEl7)
        let newEl8 = document.createElement('input')
          newEl8.id = "input_CONT_회원가입_group_allTotal_" + (num+1).toString()
          newEl8.classList.add("w3-check")
          newEl8.type = "checkbox"
          newEl8.checked = val3
          newEl7.appendChild(newEl8)
        let newEl9 = document.createElement('label')
          newEl9.htmlFor = "input_CONT_회원가입_group_allTotal_" + (num+1).toString()
          newEl9.innerHTML = " 범주" + (num+1).toString() + " 누적 총가입수 집계 (연계 우측 칸 사용)"
          newEl7.appendChild(newEl9)
        //구분선
        let newEl10 = document.createElement('hr')
          output.appendChild(newEl10)
        
        return output
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_회원가입").childElementCount
          $("#div_회원가입").appendChild(append(num))
          break
        case "remove":
          if ($("#div_회원가입").childElementCount > 1)//회원가입 항목 : 최소 1개 이상
            $("#div_회원가입").removeChild($("#div_회원가입").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_회원가입").firstElementChild)
              $("#div_회원가입").removeChild($("#div_회원가입").lastElementChild)
            cmd.wsContents.회원가입.groups.forEach((group,i) => {
              $("#div_회원가입").appendChild(
                append(i,
                  group.name.replaceAll("\n","\\n"),
                  group.items.join(","),
                  group.allTotal)
              )
            })
          }
          break
      }
      break
    
    case "강좌":
      //추가 함수 준비
      append = (num,val="",val2=false,val3="",val4="") => {
        let output = document.createElement('div')
          output.classList.add("input_CONT_강좌_group")
        //1번줄
        let newEl11 = document.createElement('p')
          output.appendChild(newEl11)
        let newEl12 = document.createElement('span')
          newEl12.classList.add("inputHeader")
          newEl12.innerHTML = "범주" + (num+1).toString() + ": "
          newEl11.appendChild(newEl12)
        let newEl13 = document.createElement('input')
          newEl13.id = "input_CONT_강좌_group_name_" + (num+1).toString()
          newEl13.classList.add("input_CONT_강좌_group_name","w3-border","long")
          newEl13.type = "text"
          newEl13.value = val
          newEl13.required = true
          newEl11.appendChild(newEl13)
        //2번 ~ 3.5번줄
        let newElCover = document.createElement('div')
          newElCover.classList.add("w3-border","w3-container","w3-light-blue")
          output.appendChild(newElCover)
        //2번줄
        let newEl21 = document.createElement('p')
          newElCover.appendChild(newEl21)
        let newEl22 = document.createElement('input')
          newEl22.id = "input_CONT_강좌_group_sub_" + (num+1).toString()
          newEl22.classList.add("w3-check")
          newEl22.type = "checkbox"
          newEl22.checked = val2
          newEl22.onchange = () => {
            switch (newEl22.checked) {
              case true:
                $("#inputP_CONT_강좌_group_subname_" + (num+1).toString()).style.display = "block"
                newElCove2.appendChild($("#inputP_CONT_강좌_group_lines_" + (num+1).toString()))
                break
              case false:
              default:
                $("#inputP_CONT_강좌_group_subname_" + (num+1).toString()).style.display = "none"
                newEl40.appendChild($("#inputP_CONT_강좌_group_lines_" + (num+1).toString()))
                break
            }
          }
          newEl21.appendChild(newEl22)
        let newEl23 = document.createElement('label')
          newEl23.htmlFor = "input_CONT_강좌_group_sub_" + (num+1).toString()
          newEl23.innerHTML = " 추가범주 사용"
          newEl21.appendChild(newEl23)
        //3번줄
        let newEl31 = document.createElement('p')
          newEl31.id = "inputP_CONT_강좌_group_subname_" + (num+1).toString()
          switch (val2) {
            case true:
              newEl31.style.display = "block"
              break
            case false:
            default:
              newEl31.style.display = "none"
              break
          }
          newElCover.appendChild(newEl31)
        let newEl32 = document.createElement('span')
          newEl32.classList.add("inputHeader")
          newEl32.innerHTML = "추가범주: "
          newEl31.appendChild(newEl32)
        let newEl33 = document.createElement('input')
          newEl33.id = "input_CONT_강좌_group_subname_" + (num+1).toString()
          newEl33.classList.add("input_CONT_강좌_group_subname","w3-border","long")
          newEl33.type = "text"
          newEl33.value = val3
          newEl33.required = true
          newEl31.appendChild(newEl33)
        //3.5번줄
        let newElCove2 = document.createElement('div')
          newElCove2.id = "inputD1_CONT_강좌_group_lines_" + (num+1).toString()
          newElCover.appendChild(newElCove2)
        //4번줄
        let newEl40 = document.createElement('div')
          newEl40.id = "inputD2_CONT_강좌_group_lines_" + (num+1).toString()
          output.appendChild(newEl40)
        let newEl41 = document.createElement('p')
          newEl41.id = "inputP_CONT_강좌_group_lines_" + (num+1).toString()
          switch (val2) {
            case true:
              newElCove2.appendChild(newEl41)
              break
            case false:
            default:
              newEl40.appendChild(newEl41)
              break
          }
        let newEl42 = document.createElement('span')
          newEl42.classList.add("inputHeader")
          newEl42.innerHTML = "입력칸수: "
          newEl41.appendChild(newEl42)
        let newEl43 = document.createElement('input')
          newEl43.id = "input_CONT_강좌_group_lines_" + (num+1).toString()
          newEl43.classList.add("input_CONT_강좌_group_lines","w3-border","long")
          newEl43.type = "text"
          newEl43.value = val4
          newEl43.required = true
          newEl43.pattern = "[0-9\,]+"
          newEl41.appendChild(newEl43)
        //구분선
        let newEl5 = document.createElement('hr')
          output.appendChild(newEl5)
        
        return output
      }
      //실행
      switch (cmd) {
        case "add":
          let num = $("#div_강좌").childElementCount
          $("#div_강좌").appendChild(append(num))
          break
        case "remove":
          if ($("#div_강좌").childElementCount > 1)//강좌 항목 : 최소 1개 이상
            $("#div_강좌").removeChild($("#div_강좌").lastElementChild)
          break
        default://cmd에 지정된 object에 있는 걸 적용
          if (typeof cmd === "object") {
            while ($("#div_강좌").firstElementChild)
              $("#div_강좌").removeChild($("#div_강좌").lastElementChild)
            cmd.wsContents.강좌.groups.forEach((group,i) => {
              if (!group.subGroup) {
                $("#div_강좌").appendChild(
                  append(i,
                    group.name.replaceAll("\n","\\n"),
                    false,
                    "",
                    group.lines)
                )
              } else {
                let nameArr = [], lineArr = []
                group.subGroup.forEach(sub => {
                  nameArr.push(sub.name)
                  lineArr.push(sub.lines)
                })
                $("#div_강좌").appendChild(
                  append(i,
                    group.name.replaceAll("\n","\\n"),
                    true,
                    nameArr.join(",").replaceAll("\n","\\n"),
                    lineArr.join(","))
                )
              }
            })
          }
          break
      }
      break
    
    default:
      break
  }
}
//추가 및 삭제 버튼
$("#input_CONT_관외대출_col_add").onclick = () => {setElement("관외대출","add")}
$("#input_CONT_관외대출_col_remove").onclick = () => {setElement("관외대출","remove")}

$("#input_CONT_이용자_col_add").onclick = () => {setElement("이용자_col","add")}
$("#input_CONT_이용자_col_remove").onclick = () => {setElement("이용자_col","remove")}

$("#input_CONT_이용자_row_add").onclick = () => {setElement("이용자_row","add")}
$("#input_CONT_이용자_row_remove").onclick = () => {setElement("이용자_row","remove")}

$("#input_CONT_소장자료_자료실_add").onclick = () => {setElement("소장자료_자료실","add")}
$("#input_CONT_소장자료_자료실_remove").onclick = () => {setElement("소장자료_자료실","remove")}

$("#input_CONT_소장자료_비도서_add").onclick = () => {setElement("소장자료_비도서","add")}
$("#input_CONT_소장자료_비도서_remove").onclick = () => {setElement("소장자료_비도서","remove")}

$("#input_CONT_소장자료_subTotal_add").onclick = () => {setElement("소장자료_subTotal","add")}
$("#input_CONT_소장자료_subTotal_remove").onclick = () => {setElement("소장자료_subTotal","remove")}

$("#input_CONT_회원가입_add").onclick = () => {setElement("회원가입","add")}
$("#input_CONT_회원가입_remove").onclick = () => {setElement("회원가입","remove")}

$("#input_CONT_강좌_add").onclick = () => {setElement("강좌","add")}
$("#input_CONT_강좌_remove").onclick = () => {setElement("강좌","remove")}

//=============================
// 3. input 반응에 따른 변형
//=============================
let changedEl
//요일 체크박스 업데이트
WEEKARR.forEach(date => {
  $$(".input_PROP_dateType_useLast").forEach(el => {
    el.onchange = () => {
      $("#div_last_" + el.value).style.display = (el.checked) ? "block" : "none"
    }
  })
})
//야간 활성화 체크박스 업데이트
$("#input_CONT_야간_enabled").onchange = () => {
  $("#div_야간").style.display = ($("#input_CONT_야간_enabled").checked) ? "block" : "none"
}
//야간 열람실 체크박스 업데이트
$("#input_CONT_야간_study").onchange = () => {
  $("#div_야간_study").style.display = ($("#input_CONT_야간_study").checked) ? "block" : "none"
}
//특성화 체크박스 업데이트
$("#input_CONT_특성화_enabled").onchange = () => {
  $("#div_특성화").style.display = ($("#input_CONT_특성화_enabled").checked) ? "block" : "none"
}

//=============================
// 4. 버튼 조작
//=============================
//시작 시
window.onload = () => {
  //저장된 데이터 불러오기 시도
  let savedData = loadContent()
  //저장된 데이터 없으면 : 기본 데이터 불러오기
  if (savedData === null) {
    setContent()
  //저장된 데이터 있으면 : 날짜를 비우고 불러오기
  } else {
    savedData.wsProperty.date = ["",""]
    setContent(savedData)
  }
  //입력값 검증 실시
  $("body").classList.add("checkInvalid")
}

//버튼 : 부가기능 (상단)
$("#settings").onclick = () => {
  $("html").scrollTo({top:$("html").scrollHeight,behavior:'smooth'})
}

//버튼 : 입력내용 반출
$("#downloadContent").onclick = () => {
  let JSONobj = JSON.stringify(getContent(), undefined, 2)
  let fileToSave = new Blob([JSONobj],{type: 'application/json'})
  saveAs(fileToSave, "data.json") 
}

//버튼 : 입력내용 반입
$("#uploadContent").onchange = (event) => {
  let reader = new FileReader()
  try {
    reader.onload = (event) => {
      try {
        let uploadObj = JSON.parse(event.target.result)
        setContent(uploadObj)
        //내용 임시저장
        saveContent()
        alert("입력내용 반입이 완료되었습니다.")
      } catch(e) {alert("* 오류 : 반입에 실패하였습니다 - 파일에 문제가 있거나, 알 수 없는 오류가 발생하였습니다.")}
    }
    reader.onerror = () => {
      alert("* 오류 : 반입에 실패하였습니다 - 파일에 문제가 있거나, 알 수 없는 오류가 발생하였습니다.")
    }
    reader.readAsText(event.target.files[0])
  } catch(e) {
    alert("* 오류 : 반입에 실패하였습니다 - 파일에 문제가 있거나, 알 수 없는 오류가 발생하였습니다.")
  }
}

//버튼 : 예시 불러오기
$("#loadExample").onclick = () => {
  let reset = confirm("예시 데이터를 불러오겠습니까?")
  if (reset === true) {
    setContent(dataObjExample)
    //내용 임시저장
    saveContent()
  }
}

//버튼 : 초기화
$("#resetContent").onclick = () => {
  let reset = confirm("현재 입력된 내용을 지우고 초기화합니다. 진행하겠습니까?")
  if (reset === true) {
    setContent()
    //내용 임시저장
    saveContent()
  }
}

//버튼 : 엑셀 출력
$("#writeExcel").onclick = () => {
  dataObj = getContent()
  writeExcel("업무일지(" + dataObj.wsProperty.library + " " + dataObj.wsProperty.date[1] + "월).xlsx")
  //내용 임시저장
  saveContent()
}

//언제나: 아무 키나 누르면: 임시저장
document.onkeydown = (e) => {
  saveContent()
}
//언제나: 마우스 좌클릭을 누르면 : 임시저장
document.onmousedown = (e) => {
  saveContent()
}