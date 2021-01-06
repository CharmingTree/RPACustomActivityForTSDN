using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPA_Controller
{
    static class GlobalConstants
    {

        public static readonly String[] RPAEQUIPCOLUMN = {"본부","TID","관리국소","fm설치위치","장치설치위치","장치대분류","장치소분류","베이","셀프","시스템번호","서비스망","사용용도","자산조직","제작사","모델명","KT자산여부","망구분","국사내여부","설치위치변경여부"};
        public static readonly String[] RPAPTNEQUIPCOLUMN = {"본부","PTN노드명(TID)","관리국소","fm설치위치","장치설치위치","장치대분류","장치소분류","베이","셀프","시스템번호","서비스망","사용용도","자산조직","제작사","모델명","KT자산여부","국사내여부","설치위치변경여부"};
        public static readonly String[] RPAUNITCOLUMN = { "설치위치", "장치명", "시스템번호", "슬롯범위", "유니트명", "유니트구분", "대역폭", "포트갯수" };
        public static readonly String[] RPACARRIERCOLUMN = {"하위설치위치","장치명","시스템","하위포트명","상위설치위치","상위장치소분류","상위장치명","상위포트명","캐리어번호","캐리어구분" };
        public static readonly String[] RPATRANSLINECOLUMN = {"하위설치위치","장치명","시스템","하위포트명","상위설치위치","캐리어번호","계위","시작타임슬롯","개수","전용회선번호","Drop연결" };
        public static readonly String[] EMSVENDORLIST = { "WR", "CO", "TF" };
        public static readonly String[] EMSRESULTCODELIST = { "Y", "N" };
    }
}
