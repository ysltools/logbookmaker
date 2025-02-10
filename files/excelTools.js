//=============================
// 1. 사전 정의
//=============================
//함수 : 숫자를 알파벳으로 변경(1을 시작점으로 잡음)
let letter = (num) => {
  return (num + 9).toString(36).toUpperCase()
}
//함수 : R/C 숫자를 A1 형식 주소 텍스트로 변경
let adr = (rI,cI,abs = "") => {
  switch (abs) {
    case "":
      return letter(cI+1) + (rI+1).toString()
    case "$r"://행 고정
      return letter(cI+1) + "$" + (rI+1).toString()
    case "$c"://열 고정
      return "$" + letter(cI+1) + (rI+1).toString()
    case "$rc"://행열 고정
    default:
      return "$" + letter(cI+1) + "$" + (rI+1).toString()
  }
}
//변수 : 시트 내용 입력용(함수 정의를 위해 사전 정의)
let sheetArr = []
//변수 : 수치 입력용 공백(테스트 시 임의 수치 출력, 이외에는 공백 출력)
let numberInput = ""

//=============================
// 2. 셀 내용 입력 함수
//=============================
let inputCell = (pos, type, value, style = {}, opt = undefined) => {
  let rI = pos[0], cI = pos[1]
  let inputArr = [type, value, style]
  //디폴트 스타일
  if (!style.al) style.al = "center"//가운데 정렬
  if (!style.bd) style.bd = "thin"//얇은 테두리
  //옵션 처리
  if (opt !== undefined) {
    inputArr.push(opt)
    //병합 옵션이 있을 경우, 병합 대상 셀 모두 비우기
    if (opt.merge !== undefined) {
      for (let r = rI;r <= rI + opt.merge.down;r++) {
        for (let c = cI;c <= cI + opt.merge.right;c++) {
          if (r !== rI || c !== cI) sheetArr[r][c] = undefined
        }
      }
    }
  }
  sheetArr[rI][cI] = inputArr
}

//=============================
// 3. 엑셀 생성
//=============================
//엑셀 출력 개시
let writeExcel = async (outputName,test = "") => {
  //테스트 시 테스트용 임의 수치 출력
  if (test === "test") numberInput = Math.floor(Math.random() * 5000) + 1000
  //콘솔만 출력할 경우, 타이머 시작
  if (outputName === "") console.time("writeExcel")

  //변수 갱신
  PROP = dataObj.wsProperty
  CONT = dataObj.wsContents

  //워크북 생성
  const workbook = new ExcelJS.Workbook()

  //시트 생성
  let sheets = ["이월"]
  for (let i = 1;i <= 31;i++) {
    sheets.push(i.toString() + "일")
  }
  sheets.forEach((s,i) => {
    workbook.addWorksheet(s)
  })
  //"1일" 시트 활성화
  workbook.views.push({activeTab:1})

  //시트 공용 뼈대 준비
  let frLen = {//temp_Frame_Width : 각 프레임별 길이 계산
    "관외대출":CONT.관외대출.cols.length,
    "이용자r":CONT.이용자.rows.length,
    "이용자c":CONT.이용자.cols.length,
    "소장실":CONT.소장자료.자료실.length,
    "비도서":CONT.소장자료.비도서.length,
    "회원가입":0,
    "강좌":0,
    "기타":CONT.기타.lines
  }
    //회원가입, 강좌는 별도로 합쳐서 길이 계산
    CONT.회원가입.groups.forEach(x => {frLen.회원가입 += x.items.length})
    CONT.강좌.groups.forEach(x => {frLen.강좌 += x.lines})
  let frame = {
    "제목":{pos:{s:{
        r: 0,
        c: 0
      },e:{
        r: 2,
        c: 15
      }}},
    "관외대출":{size:{w:16,h:frLen.관외대출+4},
      pos:{s:{
        r: 3,
        c: 0
      },e:{
        r: frLen.관외대출 + 6,
        c: 15
      }}},
    "이용자":{size:{w:frLen.이용자r*CONT.이용자.wide[0]+2*CONT.이용자.wide[1],h:frLen.이용자c+5},
      pos:{s:{
        r: frLen.관외대출 + 7,
        c: 0
      },e:{
        r: frLen.관외대출 + frLen.이용자c + 11,
        c: frLen.이용자r*CONT.이용자.wide[0] + 2*CONT.이용자.wide[1] - 1
      }}},
    "야간":{size:{w:4,h:frLen.이용자c+5},
      pos:{s:{
        r: frLen.관외대출 + 7,
        c: 12 - (2 * CONT.특성화.enabled)
      },e:{
        r: frLen.관외대출 + frLen.이용자c + 11,
        c: 15 - (2 * CONT.특성화.enabled)
      }}},
    "특성화":{size:{w:2,h:frLen.이용자c+5},pos:{s:{
        r: frLen.관외대출 + 7,
        c: 14
      },e:{
        r: frLen.관외대출 + frLen.이용자c + 11,
        c: 15
      }}},
    "소장자료":{size:{w:16,h:Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1)+2},pos:{s:{
        r: frLen.관외대출 + frLen.이용자c + 12,
        c: 0
      },e:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + 13,
        c: 15
      }}},
    "회원가입":{pos:{s:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + 14,
        c: 0
      },e:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + frLen.회원가입 + 15,
        c: 15
      }}},
    "강좌":{pos:{s:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + frLen.회원가입 + 16,
        c: 0
      },e:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + frLen.회원가입 + frLen.강좌 + 17,
        c: 15
      }}},
    "기타":{pos:{s:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + frLen.회원가입 + frLen.강좌 + 18,
        c: 0
      },e:{
        r: frLen.관외대출 + frLen.이용자c + Math.max(frLen.소장실+1,frLen.비도서,CONT.소장자료.subTotals.length*2+1) + frLen.회원가입 + frLen.강좌 + frLen.기타 + 18,
        c: 15
      }}},
  }

  //시트 입력 실시
  workbook.worksheets.forEach((worksheet, sheetNum) => {
    //이전 시트명 기억
    let preSheet = (sheetNum === 0) ? undefined : 
      ((sheetNum === 1) ? "이월" : (sheetNum-1).toString() + "일")

    //개별 시트 내용 입력 준비
    sheetArr = Array.from(
      {length:frame.기타.pos.e.r + 1},
      () => Array(frame.기타.pos.e.c + 1).fill(undefined)
    )

    //입력 예정 구역 검게 칠하기
    sheetArr.forEach((r,i) => {
      r.forEach((c,j) => {
        inputCell([i,j],"s","",{bg:"darkGray",bd:"none",ftp:"number"})
      })
    })

//================== 시트 입력 구분선 ==================
//1. 제목 입력
inputCell([0,0],"s", PROP.library + " 업무일지",
  {bd:"none",fsz:"title",fst:"bold"}, {merge:{right:15,down:0}})

//2. 날짜 및 부가정보
let dateText = ""//날짜 텍스트
let descText = ""//부가정보 텍스트
let textColor = ""//텍스트 색상
  //2-0. 이월 시트는 별도 문구 사용
  if (worksheet.name === "이월") {
    dateText = PROP.date[0].toString() + "년 " + PROP.date[1].toString() + "월 이월관리"
    descText = "'연계', '총회원수' 이월용 시트 / 이전 달 마지막 날의 '연계', '총회원수' 값을 복사 & 붙여널기 하세요"
    textColor = "FF0000"//빨강
  } else {
    //2-1. 요일 구하기
    let dateChar = WEEKARR[new Date(
      PROP.date[0] + "-" + PROP.date[1] + "-" + sheetNum.toString()
    ).getDay()]
    //2-2. 날짜 및 부가정보 텍스트 생성 및 색상 결정
    dateText = PROP.date[0].toString() + "년 "
      + PROP.date[1].toString() + "월 "
      + (sheetNum).toString() + "일 "
      + dateChar + "요일"
    if (PROP.dateType[dateChar].useLast && PROP.dateType[dateChar].last && isNaN(new Date(PROP.date[0] + "-" + PROP.date[1] + "-" + (sheetNum+7).toString()))) {
      descText = PROP.dateType[dateChar].last.desc
      textColor = PROP.dateType[dateChar].last.color
    } else {
      descText = PROP.dateType[dateChar].ordinary.desc
      textColor = PROP.dateType[dateChar].ordinary.color
    }
  }
  let cellsStyle = {bd:"none",al:"right",fsz:"large",fst:"bold",fc:textColor}
  //2-3. 날짜 및 부가정보 입력
  inputCell([1,0], "s", dateText, cellsStyle, {merge:{right:15,down:0}})
  inputCell([2,0], "s", descText, cellsStyle, {merge:{right:15,down:0}})

//3. 각 항목 입력
  //3-1. 관외대출
    //a. 헤더
    inputCell([3,0], "s", CONT.관외대출.title,
      {bg:"gray",fsz:"semiLarge",fst:"bold"},
      {merge:{right:0,down:frame.관외대출.size.h - 1}}
    )
    inputCell([3,1], "s", "구분", {bg:"gray",fsz:"middle",fst:"bold"})
    SUBJECTARR.forEach((sub,i) => {
      inputCell([3,i+2], "s", sub, {bg:"gray",fsz:"middle",fst:"bold"})
    })
    inputCell([3,12], "s", "일계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([3,13], "s", "월계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([3,14], "s", "연계", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:1,down:0}}
    )
    CONT.관외대출.cols.forEach((txt,i) => {
      inputCell([4+i,1], "s", txt, {bg:"lightGray",fsz:"middle",fst:"bold"})
    })
    inputCell([frame.관외대출.pos.e.r - 2,1], "s", "일계", {bg:"semiDarkGray",fsz:"middle",fst:"bold"})
    inputCell([frame.관외대출.pos.e.r - 1,1], "s", "월계", {bg:"semiDarkGray",fsz:"middle",fst:"bold"})
    inputCell([frame.관외대출.pos.e.r,1], "s", "연계", {bg:"semiDarkGray",fsz:"middle",fst:"bold"})
    //b. 입력칸(이월 시트 제외)
    if (worksheet.name !== "이월") {
      for (let rI = frame.관외대출.pos.s.r+1;rI <= frame.관외대출.pos.e.r-3;rI++) {
        for (let cI = frame.관외대출.pos.s.c+2;cI <= frame.관외대출.pos.e.c-4;cI++) {
          inputCell([rI,cI], "s", numberInput,
            {bg:"lightYellow",fsz:"middle",ftp:"number"})
        }
        //일계(우)
        inputCell([rI,frame.관외대출.pos.e.c-3], "f",
          "SUM(" + adr(rI,frame.관외대출.pos.s.c+2) + ":" + adr(rI,frame.관외대출.pos.e.c-4) + ")",
          {fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        //월계(우)
        inputCell([rI,frame.관외대출.pos.e.c-2], "f",
          adr(rI,frame.관외대출.pos.e.c-3) + "+'" + preSheet + "'!" + adr(rI,frame.관외대출.pos.e.c-2),
          {fsz:"middle",fst:"bold",ftp:"number"})
        //연계(우)
        inputCell([rI,frame.관외대출.pos.e.c-1], "f",
          adr(rI,frame.관외대출.pos.e.c-3) + "+'" + preSheet + "'!" + adr(rI,frame.관외대출.pos.e.c-1),
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:1,down:0}})
      }
      for (let cI = frame.관외대출.pos.s.c+2;cI <= frame.관외대출.pos.e.c-4;cI++) {
        //일계(하)
        inputCell([frame.관외대출.pos.e.r-2,cI], "f",
          "SUM(" + adr(frame.관외대출.pos.s.r+1,cI) + ":" + adr(frame.관외대출.pos.e.r-3,cI) + ")",
          {fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        //월계(하)
        inputCell([frame.관외대출.pos.e.r-1,cI], "f",
          adr(frame.관외대출.pos.e.r-2,cI) + "+'" + preSheet + "'!" + adr(frame.관외대출.pos.e.r-1,cI),
          {fsz:"middle",fst:"bold",ftp:"number"})
        //연계(하)
        inputCell([frame.관외대출.pos.e.r,cI], "f",
          adr(frame.관외대출.pos.e.r-2,cI) + "+'" + preSheet + "'!" + adr(frame.관외대출.pos.e.r,cI),
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
      }
      //일계(전체)
      inputCell([frame.관외대출.pos.e.r-2,frame.관외대출.pos.e.c-3], "f",
        //"SUM(" + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-3) + ":" + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-3) + ")",
        'IFERROR(IF(SUM(' + adr(frame.관외대출.pos.e.r-2,frame.관외대출.pos.s.c+2) + ':' +  adr(frame.관외대출.pos.e.r-2,frame.관외대출.pos.e.c-4)+ ')=SUM(' + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-3) + ':' + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-3) + '), SUM(' + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-3) + ':' + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-3) + '), "←↑합계 다름"),"←↑합계 다름")',
        {bg:"yellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
      //월계(전체)
      inputCell([frame.관외대출.pos.e.r-1,frame.관외대출.pos.e.c-2], "f",
        //"SUM(" + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-2) + ":" + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-2) + ")",
        'IFERROR(IF(SUM(' + adr(frame.관외대출.pos.e.r-1,frame.관외대출.pos.s.c+2) + ':' +  adr(frame.관외대출.pos.e.r-1,frame.관외대출.pos.e.c-4)+ ')=SUM(' + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-2) + ':' + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-2) + '), SUM(' + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-2) + ':' + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-2) + '), "←↑합계 다름"),"←↑합계 다름")',
        {fsz:"middle",fst:"bold",ftp:"number"})
      //연계(전체)
      inputCell([frame.관외대출.pos.e.r,frame.관외대출.pos.e.c-1], "f",
        //"SUM(" + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-1) + ":" + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-1) + ")",
        'IFERROR(IF(SUM(' + adr(frame.관외대출.pos.e.r,frame.관외대출.pos.s.c+2) + ':' +  adr(frame.관외대출.pos.e.r,frame.관외대출.pos.e.c-4)+ ')=SUM(' + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-1) + ':' + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-1) + '), SUM(' + adr(frame.관외대출.pos.s.r+1,frame.관외대출.pos.e.c-1) + ':' + adr(frame.관외대출.pos.e.r-3,frame.관외대출.pos.e.c-1) + '), "←↑합계 다름"),"←↑합계 다름")',
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:1,down:0}})
    //이월 시트: 연계(공란)만 입력
    } else {
      for (let rI = frame.관외대출.pos.s.r+1;rI <= frame.관외대출.pos.e.r-3;rI++) {
        //연계(우)
        inputCell([rI,frame.관외대출.pos.e.c-1], "s", numberInput,
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:1,down:0}})
      }
      for (let cI = frame.관외대출.pos.s.c+2;cI <= frame.관외대출.pos.e.c-4;cI++) {
        //연계(하)
        inputCell([frame.관외대출.pos.e.r,cI], "s", numberInput,
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
      }
      //연계(전체)
      inputCell([frame.관외대출.pos.e.r,frame.관외대출.pos.e.c-1], "s", numberInput,
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:1,down:0}})

    }
  //3-2. 이용자
    //a. 헤더
    let wide = [Number(CONT.이용자.wide[0]), Number(CONT.이용자.wide[1])]
    inputCell([frame.이용자.pos.s.r,0], "s", CONT.이용자.title, {bg:"gray",bd:"aboveThick",fsz:"semiLarge",fst:"bold"},
      {merge:{right:frame.이용자.size.w - 1,down:0}})
    inputCell([frame.이용자.pos.s.r+1,0], "s", "구분", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:wide[1]-1,down:0}})
    CONT.이용자.rows.forEach((row,i) => {
      inputCell([frame.이용자.pos.s.r+1,i*wide[0]+wide[1]], "s", row, {bg:"gray",fsz:"middle",fst:"bold"},
        {merge:{right:wide[0]-1,down:0}})
    })
    inputCell([frame.이용자.pos.s.r+1,frame.이용자.pos.e.c-(wide[1]-1)], "s", "계", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:wide[1]-1,down:0}})
    CONT.이용자.cols.forEach((col,i) => {
      inputCell([frame.이용자.pos.s.r+2+i,frame.이용자.pos.s.c], "s", col, {bg:"lightGray",fsz:"middle",fst:"bold"},
        {merge:{right:wide[1]-1,down:0}})
    })
    inputCell([frame.이용자.pos.e.r-2,0], "s", "일계", {bg:"semiDarkGray",fsz:"middle",fst:"bold"},
      {merge:{right:wide[1]-1,down:0}})
    inputCell([frame.이용자.pos.e.r-1,0], "s", "월계", {bg:"semiDarkGray",fsz:"middle",fst:"bold"},
      {merge:{right:wide[1]-1,down:0}})
    inputCell([frame.이용자.pos.e.r,0], "s", "연계", {bg:"semiDarkGray",fsz:"middle",fst:"bold"},
      {merge:{right:wide[1]-1,down:0}})
    //b. 입력칸(이월 시트 제외)
    if (worksheet.name !== "이월") {
      for (let rI = frame.이용자.pos.s.r+2;rI <= frame.이용자.pos.e.r-3;rI++) {
        for (let cI = frame.이용자.pos.s.c+wide[1];cI <= frame.이용자.pos.e.c-wide[1];cI += wide[0]) {
          inputCell([rI,cI], "s", numberInput, {bg:"lightYellow",fsz:"middle",ftp:"number"},
            {merge:{right:wide[0]-1,down:0}})
        }
        //계
        inputCell([rI,frame.이용자.pos.e.c-(wide[1]-1)], "f",
          "SUM(" + adr(rI,frame.이용자.pos.s.c+wide[1]) + ":" + adr(rI,frame.이용자.pos.e.c-wide[1]) + ")",
          {fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:wide[1]-1,down:0}})
      }
      for (let cI = frame.이용자.pos.s.c+wide[1];cI <= frame.이용자.pos.e.c-wide[1];cI += wide[0]) {
        if (cI <= frame.이용자.pos.e.c - wide[1]) {
          //일계(하) - 마지막 열 제외
          inputCell([frame.이용자.pos.e.r-2,cI], "f",
            "SUM(" + adr(frame.이용자.pos.s.r+2,cI) + ":" + adr(frame.이용자.pos.e.r-3,cI+(wide[0]-1)) + ")",
            {fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:wide[0]-1,down:0}})
          //월계(하) - 마지막 열 제외
          inputCell([frame.이용자.pos.e.r-1,cI], "f",
            adr(frame.이용자.pos.e.r-2,cI) + "+'" + preSheet + "'!" + adr(frame.이용자.pos.e.r-1,cI),
            {fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:wide[0]-1,down:0}})
          //연계(하) - 마지막 열 제외
          inputCell([frame.이용자.pos.e.r,cI], "f",
            adr(frame.이용자.pos.e.r-2,cI) + "+'" + preSheet + "'!" + adr(frame.이용자.pos.e.r,cI),
            {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:wide[0]-1,down:0}})
        }
      }
      //일계(전체)
      inputCell([frame.이용자.pos.e.r-2,frame.이용자.pos.e.c - (wide[1]-1)], "f",
        //"SUM(" + adr(frame.이용자.pos.s.r+2,frame.이용자.pos.e.c) + ":" + adr(frame.이용자.pos.e.r-3,frame.이용자.pos.e.c) + ")",
        'IFERROR(IF(SUM(' + adr(frame.이용자.pos.s.r+2,frame.이용자.pos.e.c - (wide[1]-1)) + ':' + adr(frame.이용자.pos.e.r-3,frame.이용자.pos.e.c) + ')=SUM(' + adr(frame.이용자.pos.e.r-2,frame.이용자.pos.s.c+wide[1]) + ':' + adr(frame.이용자.pos.e.r-2,frame.이용자.pos.e.c-wide[1]) + '),SUM(' + adr(frame.이용자.pos.e.r-2,frame.이용자.pos.s.c+wide[1]) + ':' + adr(frame.이용자.pos.e.r-2,frame.이용자.pos.e.c-wide[1]) + '),"←↑합계 다름"),"←↑합계 다름")',
        {bg:"yellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:wide[1]-1,down:0}})
      //월계(전체)
      inputCell([frame.이용자.pos.e.r-1,frame.이용자.pos.e.c - (wide[1]-1)], "f",
        //"SUM(" + adr(frame.이용자.pos.s.r+2,frame.이용자.pos.e.c) + ":" + adr(frame.이용자.pos.e.r-3,frame.이용자.pos.e.c) + ")",
        'IFERROR(IF(' + adr(frame.이용자.pos.e.r-2,frame.이용자.pos.e.c - (wide[1]-1)) + '+\'' + preSheet + '\'!' + adr(frame.이용자.pos.e.r-1,frame.이용자.pos.e.c) + '=SUM(' + adr(frame.이용자.pos.e.r-1,frame.이용자.pos.s.c+wide[1]) + ':' + adr(frame.이용자.pos.e.r-1,frame.이용자.pos.e.c-wide[1]) + '),SUM(' + adr(frame.이용자.pos.e.r-1,frame.이용자.pos.s.c+wide[1]) + ':' + adr(frame.이용자.pos.e.r-1,frame.이용자.pos.e.c-wide[1]) + '),"←↑합계 다름"),"←↑합계 다름")',
        {fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:wide[1]-1,down:0}})
      //연계(전체)
      inputCell([frame.이용자.pos.e.r,frame.이용자.pos.e.c - (wide[1]-1)], "f",
        //"SUM(" + adr(frame.이용자.pos.s.r+2,frame.이용자.pos.e.c) + ":" + adr(frame.이용자.pos.e.r-3,frame.이용자.pos.e.c) + ")",
        'IFERROR(IF(' + adr(frame.이용자.pos.e.r-2,frame.이용자.pos.e.c - (wide[1]-1)) + '+\'' + preSheet + '\'!' + adr(frame.이용자.pos.e.r,frame.이용자.pos.e.c) + '=SUM(' + adr(frame.이용자.pos.e.r,frame.이용자.pos.s.c+wide[1]) + ':' + adr(frame.이용자.pos.e.r,frame.이용자.pos.e.c-wide[1]) + '),SUM(' + adr(frame.이용자.pos.e.r,frame.이용자.pos.s.c+wide[1]) + ':' + adr(frame.이용자.pos.e.r,frame.이용자.pos.e.c-wide[1]) + '),"←↑합계 다름"),"←↑합계 다름")',
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:wide[1]-1,down:0}})
    //이월 시트: 연계(공란)만 입력
    } else {
      for (let cI = frame.이용자.pos.s.c+wide[1];cI <= frame.이용자.pos.e.c-wide[1];cI += wide[0]) {
        //연계(하)
        inputCell([frame.이용자.pos.e.r,cI], "s", numberInput,
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:wide[0]-1,down:0}})
      }
      inputCell([frame.이용자.pos.e.r,frame.이용자.pos.e.c-(wide[1]-1)], "s", numberInput,
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:wide[1]-1,down:0}})
    }
  //3-3. 야간이용현황(있으면)
  if (CONT.야간 && CONT.야간.enabled === 1) {
    //a. 헤더
    inputCell([frame.야간.pos.s.r,frame.야간.pos.s.c], "s", CONT.야간.title,
      {bg:"gray",bd:"aboveLeftThick",fsz:"semiLarge",fst:"bold"},
      {merge:{right:frame.야간.size.w - 1,down:0}})
    inputCell([frame.야간.pos.s.r+1,frame.야간.pos.s.c], "s", "구분", {bg:"gray",bd:"leftThick",fsz:"middle",fst:"bold"})
    inputCell([frame.야간.pos.s.r+1,frame.야간.pos.s.c+1], "s", "일계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.야간.pos.s.r+1,frame.야간.pos.s.c+2], "s", "월계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.야간.pos.s.r+1,frame.야간.pos.s.c+3], "s", "연계", {bg:"gray",fsz:"middle",fst:"bold"})
    //b. 자료이용권수
    //헤더
    inputCell([frame.야간.pos.s.r+2,frame.야간.pos.s.c], "s", "자료이용권수", {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"})
    //입력
    if (worksheet.name !== "이월") {
      inputCell([frame.야간.pos.s.r+2,frame.야간.pos.s.c+1], "f",
        "ROUNDUP(" + adr(frame.야간.pos.s.r+5,frame.야간.pos.s.c+1) + "*4/3,0)",
        {bg:"lightYellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
      inputCell([frame.야간.pos.s.r+2,frame.야간.pos.s.c+2], "f",
        adr(frame.야간.pos.s.r+2,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+2,frame.야간.pos.s.c+2),
        {fsz:"middle",fst:"bold",ftp:"number"})
      inputCell([frame.야간.pos.s.r+2,frame.야간.pos.s.c+3], "f",
        adr(frame.야간.pos.s.r+2,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+2,frame.야간.pos.s.c+3),
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
    //이월 시트: 연계(공란)만 입력
    } else {
      inputCell([frame.야간.pos.s.r+2,frame.야간.pos.s.c+3], "s", numberInput,
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
    }
    //c. 자료실이용자, 학습실 이용자
    if (CONT.야간.study === 1) {
      //헤더
      inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c], "s", "자료실이용자", {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"})
      inputCell([frame.야간.pos.s.r+4,frame.야간.pos.s.c], "s", CONT.야간.studyName, {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"})
      //입력
      if (worksheet.name !== "이월") {
        //자료실이용자
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+1], "s", numberInput,
          {bg:"lightYellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+2], "f",
          adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+2),
          {fsz:"middle",fst:"bold",ftp:"number"})
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+3], "f",
          adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+3),
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
        //학습실이용자
        inputCell([frame.야간.pos.s.r+4,frame.야간.pos.s.c+1], "s", numberInput,
          {bg:"lightYellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        inputCell([frame.야간.pos.s.r+4,frame.야간.pos.s.c+2], "f",
          adr(frame.야간.pos.s.r+4,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+4,frame.야간.pos.s.c+2),
          {fsz:"middle",fst:"bold",ftp:"number"})
        inputCell([frame.야간.pos.s.r+4,frame.야간.pos.s.c+3], "f",
          adr(frame.야간.pos.s.r+4,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+4,frame.야간.pos.s.c+3),
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
      //이월 시트: 연계(공란)만 입력
      } else {
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+3], "s", numberInput,
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})        
        inputCell([frame.야간.pos.s.r+4,frame.야간.pos.s.c+3], "s", numberInput,
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
      }
    } else {
      //헤더
      inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c], "s", "자료실이용자", {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"},
        {merge:{right:0,down:1,wrapText:false}}
      )
      //입력
      if (worksheet.name !== "이월") {
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+1], "s", numberInput,
          {bg:"lightYellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:1,wrapText:false}})
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+2], "f",
          adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+2),
          {fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:1,wrapText:false}})
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+3], "f",
          adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+3,frame.야간.pos.s.c+3),
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:1,wrapText:false}})
      //이월 시트: 연계(공란)만 입력
      } else {
        inputCell([frame.야간.pos.s.r+3,frame.야간.pos.s.c+3], "s", numberInput,
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:1,wrapText:false}})
      }
    }
    //d. 관외대출권수
    //헤더
    inputCell([frame.야간.pos.s.r+5,frame.야간.pos.s.c], "s", "관외대출권수", {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"},
      {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}}
    )
    //입력
    if (worksheet.name !== "이월") {
      inputCell([frame.야간.pos.s.r+5,frame.야간.pos.s.c+1], "s", numberInput,
        {bg:"lightYellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
      inputCell([frame.야간.pos.s.r+5,frame.야간.pos.s.c+2], "f",
        adr(frame.야간.pos.s.r+5,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+5,frame.야간.pos.s.c+2),
        {fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
      inputCell([frame.야간.pos.s.r+5,frame.야간.pos.s.c+3], "f",
        adr(frame.야간.pos.s.r+5,frame.야간.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.야간.pos.s.r+5,frame.야간.pos.s.c+3),
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
    //이월 시트: 연계(공란)만 입력
    } else {
      inputCell([frame.야간.pos.s.r+5,frame.야간.pos.s.c+3], "s", numberInput,
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
    }
  }
  //3-4. 특성화도서현황(있으면)
  if (CONT.특성화 && CONT.특성화.enabled === 1) {
    //a. 헤더
    inputCell([frame.특성화.pos.s.r,frame.특성화.pos.s.c], "s",
      CONT.특성화.title + "\n(" + CONT.특성화.name + " 도서)",
      {bg:"gray",bd:"aboveLeftThick",fsz:"semiLarge",fst:"bold"},
      {merge:{right:1,down:1}})
    inputCell([frame.특성화.pos.s.r+2,frame.특성화.pos.s.c], "s", "소장권수", {bg:"lightGray",bd:"leftThick",fsz:"middle",fst:"bold"})
    inputCell([frame.특성화.pos.s.r+3,frame.특성화.pos.s.c], "s", "대출 일계", {bg:"semiDarkGray",bd:"aboveDoubleLeftThick",fsz:"middle",fst:"bold"})
    inputCell([frame.특성화.pos.s.r+4,frame.특성화.pos.s.c], "s", "대출 월계", {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"})
    inputCell([frame.특성화.pos.s.r+5,frame.특성화.pos.s.c], "s", "대출 연계", {bg:"semiDarkGray",bd:"leftThick",fsz:"middle",fst:"bold"},
      {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
    //b. 입력칸(이월 시트 제외)
    if (worksheet.name !== "이월") {
      inputCell([frame.특성화.pos.s.r+2,frame.특성화.pos.s.c+1], "s", numberInput,
        {fsz:"middle",ftp:"number"})
      inputCell([frame.특성화.pos.s.r+3,frame.특성화.pos.s.c+1], "s", numberInput,
        {bg:"lightYellow",bd:"aboveDouble",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
      inputCell([frame.특성화.pos.s.r+4,frame.특성화.pos.s.c+1], "f",
        adr(frame.특성화.pos.s.r+3,frame.특성화.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.특성화.pos.s.r+4,frame.특성화.pos.s.c+1),
        {fsz:"middle",fst:"bold",ftp:"number"})
      inputCell([frame.특성화.pos.s.r+5,frame.특성화.pos.s.c+1], "f",
        adr(frame.특성화.pos.s.r+3,frame.특성화.pos.s.c+1) + "+'" + preSheet + "'!" + adr(frame.특성화.pos.s.r+5,frame.특성화.pos.s.c+1),
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
    //이월 시트: 연계(공란)만 입력
    } else {
      inputCell([frame.특성화.pos.s.r+5,frame.특성화.pos.s.c+1], "s", numberInput,
        {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
        {merge:{right:0,down:frame.이용자.size.h-6,wrapText:false}})
    }
  }
  //3-EX. 이용자 Row에서 공란 셀 상단 : 굵은 테두리
  sheetArr[frame.이용자.pos.s.r].forEach(cell => {
    if (cell !== undefined && cell[1] === "") {
      cell[2].bd = "aboveThickOnly"
    }
  })
  //3-5. 소장자료현황
    //a. 종합
      //a-1. 종합 헤더
      inputCell([frame.소장자료.pos.s.r,0], "s", CONT.소장자료.title,
        {bg:"gray",bd:"aboveThick",fsz:"semiLarge",fst:"bold"},
        {merge:{right:0,down:frame.소장자료.size.h - 1}}
      )
      inputCell([frame.소장자료.pos.s.r,1], "s", CONT.소장자료.totalName,
        {bg:"gray",bd:"aboveThick",fsz:"middle",fst:"bold"},
        {merge:{right:0,down:1}})
      if (CONT.소장자료.subTotals.length*2 + 3< frame.소장자료.size.h) {
        inputCell([frame.소장자료.pos.s.r + CONT.소장자료.subTotals.length*2 + 3,1],
          "s", "", {bg:"gray",fsz:"middle",fst:"bold"},
          {merge:{right:0,down:frame.소장자료.size.h - CONT.소장자료.subTotals.length*2 - 4}}
        )
      }
      CONT.소장자료.subTotals.forEach((sub,i) => {
        inputCell([frame.소장자료.pos.s.r + i*2 + 3,1], "s", sub,
          {bg:"gray",fsz:"middle",fst:"bold"})
      })
      //a-2. 종합 입력(소장총계 및 각종 연계총계 관리)
      if (worksheet.name !== "이월") {
        inputCell([frame.소장자료.pos.s.r+2,frame.소장자료.pos.s.c+1], "f",
          adr(frame.소장자료.pos.s.r+2,frame.소장자료.pos.s.c+3) + "+" + adr(frame.소장자료.pos.s.r+1,frame.소장자료.pos.e.c),
          {bg:"yellow",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
      }
      CONT.소장자료.subTotals.forEach((sub,i) => {
        if (worksheet.name !== "이월") {
          if (CONT.소장자료.subTotalsRef[i] !== "none") {
            let formula = ""
            switch (CONT.소장자료.subTotalsRef[i]) {
              case "도서":
                formula = adr(frame.소장자료.pos.s.r+2,frame.소장자료.pos.s.c+3) + " - 이월!" + adr(frame.소장자료.pos.s.r + i*2 + 4,1)
                break
              case "비도서":
                formula = adr(frame.소장자료.pos.s.r+1,frame.소장자료.pos.e.c) + " - 이월!" + adr(frame.소장자료.pos.s.r + i*2 + 4,1)
                break
              case "all":
              default:
                formula = adr(frame.소장자료.pos.s.r+2,frame.소장자료.pos.s.c+1) + " - 이월!" + adr(frame.소장자료.pos.s.r + i*2 + 4,1)
                break
            }
            inputCell([frame.소장자료.pos.s.r + i*2 + 4,1], "f", formula,
              {fc:"purple",fsz:"middle",fst:"bold",ftp:"number"})
          } else {
            inputCell([frame.소장자료.pos.s.r + i*2 + 4,1], "s",
              numberInput,
              {fc:"purple",fsz:"middle",fst:"bold",ftp:"number"})
          }
        } else {
          if (CONT.소장자료.subTotalsRef[i] !== "none") {
            inputCell([frame.소장자료.pos.s.r + i*2 + 4,1], "s",
              numberInput,
              {fc:"purple",fsz:"middle",fst:"bold",ftp:"number"},
              {note:"[" + sub + "]에 이월하여 연계되는 수치를 입력하세요.\n## 예시 ##\n* 월 증가량 : 이전 달 소장권수\n* 연 증가량 : 작년 소장권수"})
          } else {
            inputCell([frame.소장자료.pos.s.r + i*2 + 4,1], "s",
              "",
              {bg:"gray"},
              {note:"[" + sub + "]은(는) 연계되는 수치가 없습니다.\n각 월 시트에서 자유롭게 값을 입력하세요."})
          }
        }
      })
    //b. 도서
      //b-1. 도서 헤더
      inputCell([frame.소장자료.pos.s.r,2], "s", "도서 현황", {bg:"skyBlue",bd:"aboveThick",fsz:"middle",fst:"bold"},
        {merge:{right:frame.소장자료.size.w-5,down:0}})
      inputCell([frame.소장자료.pos.s.r+1,2], "s", "구분", {bg:"lightGray",fsz:"middle",fst:"bold"})
      inputCell([frame.소장자료.pos.s.r+1,3], "s", "계", {bg:"lightGray",fsz:"middle",fst:"bold"})
      SUBJECTARR.forEach((sub,i) => {
        inputCell([frame.소장자료.pos.s.r+1,4+i], "s", sub, {bg:"lightGray",fsz:"middle",fst:"bold"})
      })
      inputCell([frame.소장자료.pos.s.r+2,2], "s", "계", {bg:"lightGray",fsz:"middle",fst:"bold"})
      CONT.소장자료.자료실.forEach((row,i) => {
        inputCell([frame.소장자료.pos.s.r+3+i,2], "s", row, {bg:"lightGray",fsz:"middle",fst:"bold"})
      })
      //b-2. 도서 입력칸(이월 시트 제외)
      if (worksheet.name !== "이월") {
        for (let rI = frame.소장자료.pos.s.r+3;rI <= frame.소장자료.pos.s.r + frLen.소장실+2;rI++) {
          for (let cI = frame.소장자료.pos.s.c+4;cI <= frame.소장자료.pos.e.c-2;cI++) {
            inputCell([rI,cI], "s", numberInput, {fsz:"middle",ftp:"number"})
          }
          //도서합(좌)
          inputCell([rI,frame.소장자료.pos.s.c+3], "f",
            "SUM(" + adr(rI,frame.소장자료.pos.s.c+4) + ":" + adr(rI,frame.소장자료.pos.e.c-2) + ")",
            {fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        }
        for (let cI = frame.소장자료.pos.s.c+4;cI <= frame.소장자료.pos.e.c-2;cI++) {
          //도서합(상)
          inputCell([frame.소장자료.pos.s.r+2,cI], "f",
            "SUM(" + adr(frame.소장자료.pos.s.r+3,cI) + ":" + adr(frame.소장자료.pos.e.r,cI) + ")",
            {fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        }
        //도서합(전체)
        inputCell([frame.소장자료.pos.s.r+2,frame.소장자료.pos.s.c+3], "f",
          "SUM(" + adr(frame.소장자료.pos.s.r+3,frame.소장자료.pos.s.c+3) + ":" + adr(frame.소장자료.pos.e.r,frame.소장자료.pos.s.c+3) + ")",
          {fc:"blue",fsz:"middle",fst:"bold",ftp:"number"})
        for (let rI = frame.소장자료.pos.s.r+2;rI <= frame.소장자료.pos.s.r + frLen.비도서+1;rI++) {
          inputCell([rI,15], "s", numberInput, {fsz:"middle",fst:"bold",ftp:"number"})
        }
      }
    //c. 비도서(비도서가 없으면 출력하지 않음)
      if (CONT.소장자료.비도서.length > 0) {
        //c-1. 비도서 헤더
        inputCell([frame.소장자료.pos.s.r,frame.소장자료.pos.e.c-1], "s", "비도서 현황",
          {bg:"skyBlue",bd:"aboveThick",fsz:"middle",fst:"bold"},
          {merge:{right:1,down:0}})  
        inputCell([frame.소장자료.pos.s.r+1,frame.소장자료.pos.e.c-1], "s", "계", {bg:"lightGray",bd:"underDouble",fsz:"middle",fst:"bold"})
        //c-2. 비도서 입력칸
        CONT.소장자료.비도서.forEach((row,i) => {
          switch (CONT.소장자료.비도서type[i]) {
            case "category":
              inputCell([frame.소장자료.pos.s.r+2+i,frame.소장자료.pos.e.c-1], "s", row,
                {bg:"lightGray",fsz:"middle",fst:"bold"},{merge:{right:1,down:0}})
              break
            case "item":
            default:
              inputCell([frame.소장자료.pos.s.r+2+i,frame.소장자료.pos.e.c-1], "s", row,
                {bg:"lightGray",fsz:"middle",fst:"bold"})
              if (worksheet.name !== "이월") {
                inputCell([frame.소장자료.pos.s.r+2+i,frame.소장자료.pos.e.c], "s", numberInput,
                  {fsz:"middle",ftp:"number"})
              }
              break
          }
        })
        //c-3. 비도서 합계
        if (worksheet.name !== "이월") {
          inputCell([frame.소장자료.pos.s.r+1,frame.소장자료.pos.e.c], "f",
            "SUM(" + adr(frame.소장자료.pos.s.r+2,frame.소장자료.pos.e.c) + ":" + adr(frame.소장자료.pos.s.r+frLen.비도서+1,frame.소장자료.pos.e.c) + ")",
            {bd:"underDouble",fc:"red",fsz:"middle",fst:"bold",ftp:"number"})
        }
      }
  //3-6. 회원가입
    //a. 헤더
    inputCell([frame.회원가입.pos.s.r,0], "s", CONT.회원가입.title, {bg:"gray",bd:"aboveThick",fsz:"semiLarge",fst:"bold"},
      {merge:{right:15,down:0}}
    )
    inputCell([frame.회원가입.pos.s.r+1,0], "s", "", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:1,down:0}}
    )
    inputCell([frame.회원가입.pos.s.r+1,2], "s", "구분", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:8,down:0}}
    )
    inputCell([frame.회원가입.pos.s.r+1,11], "s", "수량", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.회원가입.pos.s.r+1,12], "s", "일계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.회원가입.pos.s.r+1,13], "s", "월계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.회원가입.pos.s.r+1,14], "s", "연계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.회원가입.pos.s.r+1,15], "s", "총가입수", {bg:"gray",fsz:"middle",fst:"bold"})
    //b. 개별그룹
    let 회원가입lines = 0
    CONT.회원가입.groups.forEach((group,i) => {
      //b-1. 개별 헤더
      inputCell([frame.회원가입.pos.s.r+회원가입lines+2,0], "s", group.name, {bg:"lightGray",bd:"underThick",fsz:"middle",fst:"bold"},
        {merge:{right:1,down:group.items.length-1}})
      group.items.forEach((item,j) => {
        let style = {fsz:"middle",fst:"bold",ftp:"number"}
        if (j === group.items.length - 1) style.bd = "underThick"
        inputCell([frame.회원가입.pos.s.r+회원가입lines+j+2,2], "s", item, style,
          {merge:{right:8,down:0}})
        if (worksheet.name !== "이월") {
          inputCell([frame.회원가입.pos.s.r+회원가입lines+j+2,11], "s", numberInput, style)
        }
      })
      //b-2. 개별 입력값
      if (worksheet.name !== "이월") {
        inputCell([frame.회원가입.pos.s.r+회원가입lines+2,12], "f", 
          "SUM(" + adr(frame.회원가입.pos.s.r+회원가입lines+2,11) + ":" + adr(frame.회원가입.pos.s.r+회원가입lines+group.items.length+1,11) + ")",
          {bd:"underThick",fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:group.items.length-1}})
        inputCell([frame.회원가입.pos.s.r+회원가입lines+2,13], "f",
          adr(frame.회원가입.pos.s.r+회원가입lines+2,12) + "+'" + preSheet + "'!" + adr(frame.회원가입.pos.s.r+회원가입lines+2,13),
          {bd:"underThick",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:group.items.length-1}})
        inputCell([frame.회원가입.pos.s.r+회원가입lines+2,14], "f",
          adr(frame.회원가입.pos.s.r+회원가입lines+2,12) + "+'" + preSheet + "'!" + adr(frame.회원가입.pos.s.r+회원가입lines+2,14),
          {bd:"underThick",fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:group.items.length-1}})
        if (group.allTotal === true) {
          inputCell([frame.회원가입.pos.s.r+회원가입lines+2,15], "f",
            adr(frame.회원가입.pos.s.r+회원가입lines+2,12) + "+'" + preSheet + "'!" + adr(frame.회원가입.pos.s.r+회원가입lines+2,15),
            {bd:"underThick",fc:"purple",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:0,down:group.items.length-1}})
        }
      //이월 시트: 연계(공란)만 입력
      } else {
        inputCell([frame.회원가입.pos.s.r+회원가입lines+2,14], "s", numberInput,
          {bd:"underThick",fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
          {merge:{right:0,down:group.items.length-1}})
        if (group.allTotal === true) {
          inputCell([frame.회원가입.pos.s.r+회원가입lines+2,15], "s", numberInput,
            {bd:"underThick",fc:"purple",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:0,down:group.items.length-1}})
        }
      }
      //b-3. 라인 수 넘기기
      회원가입lines += group.items.length
    })
  //3-6. 강좌
    //a. 헤더
    inputCell([frame.강좌.pos.s.r,0], "s", CONT.강좌.title, {bg:"gray",bd:"aboveThick",fsz:"semiLarge",fst:"bold"},
      {merge:{right:15,down:0}}
    )
    inputCell([frame.강좌.pos.s.r+1,0], "s", "구분", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:1,down:0}}
    )
    inputCell([frame.강좌.pos.s.r+1,2], "s", "내용", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:5,down:0}}
    )
    inputCell([frame.강좌.pos.s.r+1,8], "s", "일시", {bg:"gray",fsz:"middle",fst:"bold"},
      {merge:{right:2,down:0}}
    )
    inputCell([frame.강좌.pos.s.r+1,11], "s", "인원", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.강좌.pos.s.r+1,12], "s", "일계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.강좌.pos.s.r+1,13], "s", "월계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.강좌.pos.s.r+1,14], "s", "연계", {bg:"gray",fsz:"middle",fst:"bold"})
    inputCell([frame.강좌.pos.s.r+1,15], "s", "비고", {bg:"gray",fsz:"middle",fst:"bold"})
    //b. 각 그룹
    let 강좌lines = 0, 강좌gbNum = 0, 강좌bg = ""
    CONT.강좌.groups.forEach((group,i) => {
      if (group.subGroup !== undefined) {
        inputCell([frame.강좌.pos.s.r+강좌lines+2,0], "s", group.name,
          {bg:"lightGray",fsz:"semiLarge",fst:"bold"},
          {merge:{right:0,down:group.lines - 1}}
        )
        let 강좌subLines = 0
        group.subGroup.forEach((sub,j) => {
          inputCell([frame.강좌.pos.s.r+강좌lines+강좌subLines+2,1], "s", sub.name,
            {bg:"lightGray",fsz:"semiLarge",fst:"bold"},
            {merge:{right:0,down:sub.lines - 1}}
          )
          /*배경색*/ if (!Number.isInteger(강좌gbNum/2)) {강좌bg = "none"} else {강좌bg = cellStyles.강좌bgs[(강좌gbNum / 2) % cellStyles.강좌bgs.length]}
          for (let rI = frame.강좌.pos.s.r+강좌lines+강좌subLines+2;rI < frame.강좌.pos.s.r+강좌lines+강좌subLines+sub.lines+2;rI++) {
            inputCell([rI,2], "s", "", {bg:강좌bg,fsz:"middle"}, {merge:{right:5,down:0}})
            inputCell([rI,8], "s", "", {bg:강좌bg,fsz:"middle"}, {merge:{right:2,down:0}})
            if (worksheet.name !== "이월") {
              inputCell([rI,11], "s", numberInput, {bg:강좌bg,fsz:"middle",fst:"bold",ftp:"number"})
            }
          }
          if (worksheet.name !== "이월") {
            inputCell([frame.강좌.pos.s.r+강좌lines+강좌subLines+2,12], "f", 
              "SUM(" + adr(frame.강좌.pos.s.r+강좌lines+강좌subLines+2,11) + ":" + adr(frame.강좌.pos.s.r+강좌lines+강좌subLines+sub.lines+1,11) + ")",
              {bg:강좌bg,fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
              {merge:{right:0,down:sub.lines-1}})
            inputCell([frame.강좌.pos.s.r+강좌lines+강좌subLines+2,13], "f",
              adr(frame.강좌.pos.s.r+강좌lines+강좌subLines+2,12) + "+'" + preSheet + "'!" + adr(frame.강좌.pos.s.r+강좌lines+강좌subLines+2,13),
              {bg:강좌bg,fsz:"middle",fst:"bold",ftp:"number"},
              {merge:{right:0,down:sub.lines-1}})
            inputCell([frame.강좌.pos.s.r+강좌lines+강좌subLines+2,14], "f",
              adr(frame.강좌.pos.s.r+강좌lines+강좌subLines+2,12) + "+'" + preSheet + "'!" + adr(frame.강좌.pos.s.r+강좌lines+강좌subLines+2,14),
              {bg:강좌bg,fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
              {merge:{right:0,down:sub.lines-1}})
            inputCell([frame.강좌.pos.s.r+강좌lines+강좌subLines+2,15], "s", "",
              {bg:강좌bg,fsz:"middle",fst:"bold"},
              {merge:{right:0,down:sub.lines-1}})
          //이월 시트: 연계(공란)만 입력
          } else {
            inputCell([frame.강좌.pos.s.r+강좌lines+강좌subLines+2,14], "s", numberInput,
              {bg:강좌bg,fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
              {merge:{right:0,down:sub.lines-1}})
          }
          강좌gbNum += 1
          강좌subLines += sub.lines
        })
      } else {
        inputCell([frame.강좌.pos.s.r+강좌lines+2,0], "s", group.name,
          {bg:"lightGray",fsz:"semiLarge",fst:"bold"},
          {merge:{right:1,down:group.lines - 1}}
        )
        /*배경색*/ if (!Number.isInteger(강좌gbNum/2)) {강좌bg = "none"} else {강좌bg = cellStyles.강좌bgs[(강좌gbNum / 2) % cellStyles.강좌bgs.length]}
        for (let rI = frame.강좌.pos.s.r+강좌lines+2;rI < frame.강좌.pos.s.r+강좌lines+group.lines+2;rI++) {
          inputCell([rI,2], "s", "", {bg:강좌bg,fsz:"middle"}, {merge:{right:5,down:0}})
          inputCell([rI,8], "s", "", {bg:강좌bg,fsz:"middle"}, {merge:{right:2,down:0}})
          if (worksheet.name !== "이월") {
            inputCell([rI,11], "s", numberInput, {bg:강좌bg,fsz:"middle",fst:"bold",ftp:"number"})
          }
        }
        if (worksheet.name !== "이월") {
          inputCell([frame.강좌.pos.s.r+강좌lines+2,12], "f", 
            "SUM(" + adr(frame.강좌.pos.s.r+강좌lines+2,11) + ":" + adr(frame.강좌.pos.s.r+강좌lines+group.lines+1,11) + ")",
            {bg:강좌bg,fc:"red",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:0,down:group.lines-1}})
          inputCell([frame.강좌.pos.s.r+강좌lines+2,13], "f",
            adr(frame.강좌.pos.s.r+강좌lines+2,12) + "+'" + preSheet + "'!" + adr(frame.강좌.pos.s.r+강좌lines+2,13),
            {bg:강좌bg,fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:0,down:group.lines-1}})
          inputCell([frame.강좌.pos.s.r+강좌lines+2,14], "f",
            adr(frame.강좌.pos.s.r+강좌lines+2,12) + "+'" + preSheet + "'!" + adr(frame.강좌.pos.s.r+강좌lines+2,14),
            {bg:강좌bg,fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:0,down:group.lines-1}})
          inputCell([frame.강좌.pos.s.r+강좌lines+2,15], "s", "",
            {bg:강좌bg,fsz:"middle",fst:"bold"},
            {merge:{right:0,down:group.lines-1}})
        //이월 시트: 연계(공란)만 입력
        } else {
          inputCell([frame.강좌.pos.s.r+강좌lines+2,14], "s", numberInput,
            {bg:강좌bg,fc:"blue",fsz:"middle",fst:"bold",ftp:"number"},
            {merge:{right:0,down:group.lines-1}})
        }
        강좌gbNum += 1
      }
      강좌lines += group.lines
    })
  //3-6. 기타(특기 사항)
    //a. 헤더
    inputCell([frame.기타.pos.s.r,0], "s", CONT.기타.title, {bg:"gray",bd:"aboveThick",fsz:"semiLarge",fst:"bold"},
      {merge:{right:15,down:0}}
    )
    //b. 입력칸(이월 시트 제외)
    if (worksheet.name !== "이월") {
      for (let i = 0;i <= CONT.기타.lines - 1;i++) {
        if (i < CONT.기타.lines - 1) {
          inputCell([frame.기타.pos.s.r+1+i,0], "s", "", {bd:"underWhite",al:"left",fsz:"semiLarge",fst:"bold"},
            {merge:{right:15,down:0}}
          )
        } else {
          inputCell([frame.기타.pos.s.r+1+i,0], "s", "", {bd:"aboveWhite",al:"left",fsz:"semiLarge",fst:"bold"},
            {merge:{right:15,down:0}}
          )
        }
      }
    }

//================== 시트 입력 구분선 ==================
    
    //시트 내용 적용
    sheetArr.forEach((r,rI) => {
      r.forEach((c,cI) => {
        //※ 해당 셀 입력용 데이터가 없으면 넘기기
        if (sheetArr[rI][cI] === undefined) return
        //셀 입력용 데이터 불러오기 / cellData : [타입, 값, 스타일, 옵션]
        let cellData = sheetArr[rI][cI]
        //셀 정의
        let cell = worksheet.getCell(letter(cI+1) + (rI+1).toString())
        //셀 값(공란이 아닐 경우에만 적용)
        if (cellData[1] !== "") {
          switch (cellData[0]) {
            case "f"://0번 값이 "f"면 함수 출력
              cell.value = {formula: cellData[1]}
              break
            case "s"://0번 값이 "s"면 단순 문자 출력(숫자도 포함)
            default:
              cell.value = cellData[1]
              break
          }
        }
        //기본 폰트 적용
        cell.font = {}
        cell.font.name = PROP.fontName
        //셀 스타일(cellStyles에서 찾아서 적용)
        for (let styleKey in cellData[2]) {
          switch (styleKey) {
            //폰트 색상
            case "fc":
              cell.font.color = {}
              let colorTxt = ""
              if (!cellStyles[styleKey][cellData[2][styleKey]])
                colorTxt = cellData[2][styleKey]
              else colorTxt = cellStyles[styleKey][cellData[2][styleKey]]
              cell.font.color.argb = colorTxt
              break
            //폰트 크기 : 문구를 대응하는 수치로 변환(미대응 시 middle 적용)
            case "fsz":
              let fszText = cellData[2][styleKey]
              if (PROP.fontSize[fszText] !== undefined) {
                cell.font.size = PROP.fontSize[fszText]
              } else {
                cell.font.size = PROP.fontSize.middle
              }
              break
            //그 외 스타일
            default:
              let styleObj = cellStyles[styleKey][cellData[2][styleKey]]
              let styleType = Object.keys(styleObj)[0]
              if (typeof styleObj[styleType] === "string") {
                if (!cell[styleType]) cell[styleType] = ""
                cell[styleType] = styleObj[styleType]
              } else {
                if (!cell[styleType]) cell[styleType] = {}
                for (let subKey in styleObj[styleType]) {
                  cell[styleType][subKey] = styleObj[styleType][subKey]
                }
              }
              break
          }
        }
        //별도 옵션 처리
        if (cellData[3] !== undefined) {
          let opt = cellData[3]
          //옵션 : 병합(right : 가로 → / down : 세로 ↓)
          if (opt.merge !== undefined) {
            //병합셀 속성 변경 / '셀에 맞춤' → '텍스트 줄 바꿈'
            if (opt.merge.down !== 0 && opt.merge.wrapText !== false) {//down이 1 이상이거나  wrapText가 false가 아니면 적용
              if (!cell.alignment) cell.alignment = {}
              cell.alignment.shrinkToFit = false
              cell.alignment.wrapText = true
            }
            //병합 실시
            worksheet.mergeCells(
              rI+1, cI+1,//시작행, 시작열
              rI+1+opt.merge.down, cI+1+opt.merge.right//종료행, 종료열
            )
          }
          //옵션 : 메모 / 옵션 내용을 그대로 적용
          if (opt.note !== undefined) {
            cell.note = {texts:[{font:{size:PROP.fontSize.small},text:opt.note}]}
          }
        }
      })
    })

    //시트 행, 열 규격, 확대
      //행 규격
      worksheet.eachRow((row,i) => {
        switch (i) {
          case 1:
            row.height = PROP.rowHeight.first
            break
          default:
            row.height = PROP.rowHeight.other
            break
        }
      })
      //열 규격
      worksheet.columns.forEach((col) => {
        col.width = PROP.colWidth
      })
      //확대
      worksheet.views = [{zoomScale: PROP.zoom}]
      worksheet.properties.defaultColWidth = PROP.colWidth//오류 방지용
      worksheet.properties.defaultRowHeight = PROP.rowHeight.other//오류 방지용

  })

  //엑셀 파일 최종 작성
  if (outputName === "") {
    //테스트 : 엑셀 작성하지 않고 콘솔에서만 확인, 타이머 출력
    console.log(workbook)
    console.timeEnd("writeExcel")
    return
  }
  let wbout = await workbook.xlsx.writeBuffer()
  //엑셀 파일 다운로드
  saveAs(new Blob([wbout],{type:"application/octet-stream"}), outputName)
}