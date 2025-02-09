//엑셀 각 셀 스타일 관리
let cellStyles = {
  //배경색
  bg:{
    none:{//없음
      fill:{type:"pattern",pattern:"solid",fgColor:{}}
    },
    gray:{//회색
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"BFBFBF"}}
    },
    lightGray:{//연한 회색
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"F2F2F2"}}
    },
    darkGray:{//진한 회색
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"808080"}}
      //fill:{type:"pattern",pattern:"solid",fgColor:{argb:"595959"}}/*더 진한 색*/
    },
    semiDarkGray:{//약간 진한 회색
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"D9D9D9"}}
      //fill:{type:"pattern",pattern:"solid",fgColor:{argb:"595959"}}/*더 진한 색*/
    },
    yellow:{//노랑
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"FFFF00"}}
    },
    lightYellow:{//연한 노랑
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"FFF2CC"}}
    },
    skyBlue:{//하늘색
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"DDEBF7"}}
    },
    lightGreen:{//연한녹색
      fill:{type:"pattern",pattern:"solid",fgColor:{argb:"E2EFDA"}}
    }
  },
  //테두리
  bd:{
    thin:{//디폴트값(inputCell 함수에서 처리)
      border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
    },
    none:{//테두리 없음
      border:{}
    },
    aboveDouble:{//"특성화" 카테고리용 1
      border:{top:{style:"double"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
    },
    aboveDoubleLeftThick:{//"특성화" 카테고리용 2
      border:{top:{style:"double"},bottom:{style:"thin"},left:{style:"medium"},right:{style:"thin"}}
    },
    aboveThick:{
      border:{top:{style:"medium"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
    },
    aboveThickOnly:{
      border:{top:{style:"medium"}}
    },
    aboveLeftThick:{
      border:{top:{style:"medium"},bottom:{style:"thin"},left:{style:"medium"},right:{style:"thin"}}
    },
    leftThick:{
      border:{top:{style:"h"},bottom:{style:"thin"},left:{style:"medium"},right:{style:"thin"}}
    },
    underThick:{
      border:{top:{style:"thin"},bottom:{style:"medium"},left:{style:"thin"},right:{style:"thin"}}
    },
    underWhite:{//"기타" 카테고리용 1
      border:{bottom:{style:"thin",color:{argb:"ffffff"}},left:{style:"thin"},right:{style:"thin"}}
    },
    aboveWhite:{//"기타" 카테고리용 2
      border:{top:{style:"thin",color:{argb:"ffffff"}},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
    }
  },
  //글자 정렬
  al:{
    center:{//디폴트값(inputCell 함수에서 처리)
      alignment:{vertical:"middle",horizontal:"center",shrinkToFit:true}
    },
    right:{
      alignment:{vertical:"middle",horizontal:"right",shrinkToFit:true}
    },
    left:{
      alignment:{vertical:"middle",horizontal:"left",shrinkToFit:true}
    }
  },
  //폰트 색상(HEX값만 입력, 스타일 적용 시 옵션에서 별도로 처리)
  fc:{
    //검정 : 미지정 시 디폴트값
    red:"FF0000",
    blue:"0000FF",
    purple:"A020F0"
  },
  //폰트 크기(입력값 그대로 스타일 적용 시 처리)
  fsz:{},//large, small 등 / 
  //폰트 스타일
  fst:{
    bold:{
      font:{bold:"true"}
    }
  },
  //입력 타입
  ftp:{
    number:{
      numFmt:"#,##0"
    }
  },
  //강좌 배경색: 번갈아가면 적용됨
  강좌bgs:["lightYellow","skyBlue","lightGreen"]
}