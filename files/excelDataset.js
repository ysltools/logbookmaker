let dataObj = 
{
  wsProperty:{
    library:"○○도서관",
    date:["",""],
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
    dateType:{
      "월":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },last:{desc:"",color:""},
      useLast:false},
      "화":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },last:{desc:"",color:""},
      useLast:false},
      "수":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },last:{desc:"",color:""},
      useLast:false},
      "목":{ordinary:{
        desc:"자료실 및 학습실 야간 운영",
        color:"0000FF"
      },last:{desc:"",color:""},
      useLast:false},
      "금":{ordinary:{
        desc:"자료실 휴실, 학습실 주간 운영",
        color:"0000FF"
      },last:{
        desc:"전체 휴관일",
        color:"FF0000"
      },useLast:true},
      "토":{ordinary:{
        desc:"자료실 주간 운영, 학습실 야간 운영",
        color:"0000FF"
      },last:{desc:"",color:""},
      useLast:false},
      "일":{ordinary:{
        desc:"자료실 주간 운영, 학습실 야간 운영",
        color:"0000FF"
      },last:{desc:"",color:""},
      useLast:false},
    }
  },
  wsContents:{
    "관외대출":{
      title:"📓\n관외대출\n현황",
      cols:["종합자료실","어린이자료실","동네서점"]
    },
    "이용자":{
      title:"👪 이용자 현황",
      maxRow:14,
      rows:["종합자료실","어린이자료실","동네서점","강좌 및 행사"],
      cols:["영유아","학생","일반"],
      wide:"11"//이용자 현황 확장 키워드
    },
    "야간":{
      title:"🌙 야간 이용 현황",
      enabled:0,
      study:1,
      studyName:"학습실이용자"
    },
    "특성화":{
      title:"☺ 특성화자료 현황",
      name:"향토",
      enabled:0
    },
    "소장자료":{
      title:"📕\n소장자료\n현황",
      자료실:["종합자료실","어린이자료실","보존서고"],
      비도서:["시청각자료","DVD","부록자료","CD","TAPE","딸림책"],
      비도서type:["category","item","category","item","item","item"],
      totalName:"📚\n소장총계",
      subTotals:["年 도서증가"],
      subTotalsRef:["도서"]//도서, 비도서, all, none 가능
    },
    "회원가입":{
      title:"📝 독서회원 가입 현황",
      groups:[
        {
          name:"회원 가입",
          items:["영유아 및 학생","일반"],
          allTotal:true
        }
      ]
    },
    "강좌":{
      title:"🖍 강좌 및 행사 운영 현황",
      lines:14,
      groups:[
        {
          name:"독서\n진흥\n강좌\n및\n행사",
          lines:11,
          subGroup:[
            {name:"문화강좌",lines:6},
            {name:"독서회",lines:3},
            {name:"기타\n독서진흥\n프로그램",lines:3}
          ]
        },
        {
          name:"도서관 견학\n(이용자 교육)",
          lines:3
        }
      ]
    },
    "기타":{
      title:"📌 특기사항",
      lines:4
    }
  }
}