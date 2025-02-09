let dataObj = 
{
  wsProperty:{
    fontName:"맑은 고딕",
    fontSize:{
      title:30,
      large:20,
      semiLarge:18,
      middle:16,
      small:14
    },
    rowHeight:{
      first:50,
      other:30
    },
    colWidth:17,
    zoom:60,
    library:"서창도서관",
    date:["2025","2"],
    강좌bgs:["lightYellow","skyBlue","lightGreen"],//강좌 배경색은 여기서 번갈아가면 적용됨
    dateType:{
      "월":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },useLast:false},
      "화":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },useLast:false},
      "수":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },useLast:false},
      "목":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },useLast:false},
      "금":{ordinary:{
        desc:"자료실 휴실, 학습실 주간 운영",
        color:"A020F0"
      },last:{
        desc:"전체 휴관일",
        color:"FF0000"
      },useLast:true},
      "토":{ordinary:{
        desc:"자료실 주간 운영, 학습실 야간 운영",
        color:"0000FF"
      },useLast:false},
      "일":{ordinary:{
        desc:"자료실 주간 운영, 학습실 야간 운영",
        color:"0000FF"
      },useLast:false},
    }
  },
  wsContents:{
    "관외대출":{
      title:"📓\n관외대출\n현황",
      cols:["자료실1","자료실2","자료실3","자료실4","자료실5","자료실6","자료실7","자료실8"]
    },
    "이용자":{
      title:"👪 이용자 현황",
      maxRow:10,
      rows:["분류1","분류2","분류3","분류4","분류5","분류6","분류7"],
      cols:["영유아","학생","일반"]
    },
    "야간":{
      title:"🌙 야간 이용 현황",
      enabled:1,
      study:1,
      studyName:"학습실이용자"
    },
    "특성화":{
      title:"☺ 특성화자료 현황",
      name:"웹툰",
      enabled:1
    },
    "소장자료":{
      title:"📕\n소장자료\n현황",
      자료실:["자료실1","자료실2","자료실3","자료실4","자료실5","자료실6"],
      비도서:["시청각 자료","부록1","부록 자료","부록2","부록3","부록4"],
      비도서type:["category","item","category","item","item","item"],
      totalName:"📚\n소장총계",
      subTotals:["月 월간증가","年 연간증가","🌐 전자저널"],
      subTotalsRef:["도서","all","none"]
    },
    "회원가입":{
      title:"📝 독서회원 가입 현황",
      groups:[
        {
          name:"회원 가입",
          items:["영유아","학생","성인","노년층"],
          allTotal:true
        },
        {
          name:"회원증 발급",
          items:["최초발급","재발급"],
          allTotal:false
        },
        {
          name:"소셜 가입",
          items:["카카오친구","인스타그램","카페"],
          allTotal:true
        }
      ]
    },
    "강좌":{
      title:"🖍 강좌 및 행사 운영 현황",
      lines:15,
      groups:[
        {
          name:"독서\n진흥\n강좌\n및\n행사",
          lines:11,
          subGroup:[
            {name:"문화강좌",lines:6},
            {name:"독서회",lines:2},
            {name:"기타\n독서진흥\n프로그램",lines:3}
          ]
        },
        {
          name:"도서관 견학\n(자유견학 및\n이용자 교육)",
          lines:3
        },
        {
          name:"독서동아리 강의실 지원",
          lines:1
        }
      ]
    },
    "기타":{
      title:"📌 특기 사항",
      lines:12
    }
  }
}