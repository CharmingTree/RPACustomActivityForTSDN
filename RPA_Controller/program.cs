using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Activities;
using System.ComponentModel;
using System.Diagnostics;
using outlook = Microsoft.Office.Interop.Outlook;
using excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Text.RegularExpressions;

using Newtonsoft.Json;
using Js = Newtonsoft.Json.Linq;


namespace RPA_Controller
{
    namespace Common
    {
        namespace FileControll
        {
            public class ExcelFileTypeSetup : CodeActivity
            {
                [DisplayName("FilePath"), Category("Input")]
                [RequiredArgument]
                [Description("파일 경로를 입력해주세요")]
                public InArgument<String> FilePath { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    String path = FilePath.Get(context);
                    configureWorkbook(path);
                }

                public void logError(Exception e)
                {
                    Console.WriteLine("<<E>> ExcelFileTypeSetup - 예외 / 타입 : {0}, 메세지 : \n{1}", e.GetType().ToString(), e.Message);
                }

                public void configureWorkbook(String workbookPath)
                {
                    try
                    {
                        configureTextCell(workbookPath);
                    }
                    catch (Exception e)
                    {
                        logError(e);
                    }
                }

                private void configureTextCell(String path)
                {
                    excel.Application app = new excel.Application();
                    // excel Display off 
                    app.Visible = false; // 백그라운드로 처리
                    app.DisplayAlerts = false; // 엑셀 경고창 표시안함

                    excel.Workbook workbook = app.Workbooks.Open(Filename: @path);

                    /* 시트 OPEN */
                    excel.Worksheet worksheet = workbook.Worksheets.Item[1];


                    // 텍스트형 셀 설정 
                    excel.Range TextCell = worksheet.UsedRange;
                    //excel.Range TextCell = worksheet.Range["A1", "Z99"];
                    TextCell.NumberFormat = "@";

                    workbook.Save();
                    workbook.Close();
                    app.Quit();
                }

            }

            /* ###############################################################
            파일 복사 & 붙여넣기 기능 수행
           ###############################################################*/
            [Designer(typeof(FileCopyActivityDesigner))]
            public class FileCopyActivity : CodeActivity
            {
                [DisplayName("filePath"), Category("Input")]
                [Description("복사 대상 파일 경로")]
                [RequiredArgument]
                public InArgument<String> filePath { get; set; }

                [DisplayName("targetFilePath"), Category("Input")]
                [Description("복사되는 파일 경로")]
                [RequiredArgument]
                public InArgument<String> targetFilePath { get; set; }

                [DisplayName("coverYN"), Category("Input")]
                [Description("기존 파일 덮어씌울지 여부")]
                [RequiredArgument]
                public InArgument<bool> coverYN { get; set; }

                [Category("Output")]
                [Description("복사 성공 여부를 반환합니다.")]
                public OutArgument<bool> successYN { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    String _filepath = @filePath.Get(context);
                    String _targetFilePath = @targetFilePath.Get(context);
                    bool _coverYN = coverYN.Get(context);
                    bool _successYN = false;

                    FileInfo fileObj = new FileInfo(_filepath);

                    if (fileObj.Exists)
                    {
                        fileObj.CopyTo(_targetFilePath, _coverYN);
                        FileInfo targetFileObj = new FileInfo(_targetFilePath);
                        if (targetFileObj.Exists)
                        {
                            Console.WriteLine("RPA_Controller : FileCopyActivity Logger.info - 성공적으로 파일이 복사되었습니다. 복사 대상 : {0} / 복사된 파일 : {1}", _filepath, _targetFilePath);
                            _successYN = true;
                        }
                        else
                        {
                            Console.WriteLine("RPA_Controller : FileCopyActivity Logger.info - 복사된 파일이 존재하지 않습니다.");
                            _successYN = false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("RPA_Controller : FileCopyActivity Logger.info - 복사 대상 파일이 존재하지 않습니다.");
                    }


                    successYN.Set(context, _successYN);
                }
            }
        } // end of FileControll

        namespace Assgin
        {
            public class ClearTable : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<DataSet> dataSet { get; set; }

                [Category("Input")]
                public InArgument<String> tableName { get; set; }

                [Category("Output")]
                public OutArgument<DataSet> resultSet { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    dataSet.Get(context).Tables[tableName.Get(context)].Clear();
                    resultSet.Set(context, dataSet.Get(context));
                }
            }

            /*
             * DataSet 객체에 Table을 추가하는 기능.
             */
            public class AddTable : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<DataSet> dataset { get; set; }
                [Category("Input"), RequiredArgument]
                public InArgument<DataTable> datatable { get; set; }

                [Category("Option"), Description("같은 테이블 이름을 가진 경우 덮어쓸지 여부를 결정합니다.")]
                public InArgument<Boolean> overwriteYN { get; set; }

                [Category("Output")]
                public OutArgument<DataSet> resultSet { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Boolean overwriteYN = this.overwriteYN.Get(context);

                    Console.WriteLine("<<C>> AddTable - Start");
                    if (dataset.Get(context) == null)
                    {
                        Console.WriteLine("<<C>> AddTable - DataSet is null");
                    }
                    else if (datatable.Get(context) == null || datatable.Get(context).Rows.Count == 0)
                    {
                        Console.WriteLine("<<C>> AddTable - DataTable is null");
                        resultSet.Set(context, dataset.Get(context));
                    }
                    else
                    {
                        if (overwriteYN)
                        {
                            if (dataset.Get(context).Tables.Contains(datatable.Get(context).TableName))
                            {
                                Console.WriteLine("DataSet에 이미 존재하는 테이블이지만 덮어 씌웁니다.");
                                dataset.Get(context).Tables.Remove(datatable.Get(context).TableName);
                                dataset.Get(context).Tables.Add(datatable.Get(context));
                                resultSet.Set(context, dataset.Get(context));
                            }
                            else
                            {
                                dataset.Get(context).Tables.Add(datatable.Get(context));
                                resultSet.Set(context, dataset.Get(context));
                            }
                        }
                        else
                        {
                            if (dataset.Get(context).Tables.Contains(datatable.Get(context).TableName))
                            {
                                Console.WriteLine("<<E>> AddTable - DataSet에 이미 존재하는 테이블 입니다.");
                                resultSet.Set(context, dataset.Get(context));
                            }
                            else
                            {
                                dataset.Get(context).Tables.Add(datatable.Get(context));
                                resultSet.Set(context, dataset.Get(context));
                            }
                        }
                    }
                    Console.WriteLine("<<C>> AddTable - End");

                }
            }

            /*
             * Dictionary에 요소 하나를 추가합니다.
             */
            public class DictionaryElementAdd : CodeActivity
            {
                [Category("Input")]
                public InOutArgument<Dictionary<String, String>> pDictionary { get; set; }
                [Category("Input")]
                public InArgument<String> pKey { get; set; }
                [Category("Input")]
                public InArgument<String> pValue { get; set; }

                [Category("output")]
                public OutArgument<Dictionary<String,String>> result { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    if (pDictionary.Get(context) == null || pKey.Get(context) == null || pKey.Get(context).Equals("") || pValue.Get(context) == null)
                    {
                        Console.WriteLine("<<E>> DictionaryElementAdd - parameter is null");
                    }
                    else
                    {
                        pDictionary.Get(context).Add(pKey.Get(context), pValue.Get(context));

                        result.Set(context, pDictionary.Get(context));
                    }
                }

            }
        }
    
        namespace StringFunction
        {
            public class RegMatch : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<String> input { get; set; }
                [Category("Input"), RequiredArgument]
                public InArgument<String> regularExp { get; set; }

                [Category("Output")]
                public OutArgument<String> output { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Boolean check = true;
                    String tmpStr = input.Get(context);
                    String tmpRegularExp = regularExp.Get(context);

                    if (tmpStr != null && tmpRegularExp != null)
                    {
                        Regex reg = new Regex(tmpRegularExp);
                        MatchCollection result = reg.Matches(tmpStr);

                        if (result != null && result.Count > 0)
                        {
                            foreach (Match m in result)
                            {
                                output.Set(context, m.Groups[0].ToString());
                                break;
                            }
                        }
                        else { check = false; }
                    }
                    else
                    {
                        check = false;
                    }

                    if (!check)
                    {
                        output.Set(context, "null");
                    }
                }
            }
        }
    } // end of Common

    namespace FmMismatch
    {
        namespace Excel
        {
            public class PTNSystemReRegHist : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<DataRow> reRegRow { get; set; }

                [Category("Input"), RequiredArgument]
                public InArgument<String> regResult { get; set; }

                [Category("Input")]
                public InArgument<String> regCode { get; set; }

                [Category("Input")]
                public InArgument<String> regCause { get; set; }

                [Category("regHistDt"), RequiredArgument]
                public InOutArgument<DataTable> reRegHistDt { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    UpdateHistTable(reRegRow.Get(context), reRegHistDt.Get(context), regResult.Get(context), regCode.Get(context), regCause.Get(context));
                }

                private void UpdateHistTable(DataRow row, DataTable table, String regResult, String regCode, String regCause)
                {
                    try
                    {
                        if (row == null)
                        {
                            Console.WriteLine("null입니다");
                        }

                        if (table == null || table.Rows.Count == 0)
                        {
                            Console.WriteLine("null입니다2");
                        }
                        foreach (DataRow i in table.Rows)
                        {
                            if (row["장치설치위치"].ToString().Equals(i["장치설치위치"].ToString()) &&
                               row["베이"].ToString().Equals(i["베이"].ToString()) &&
                               row["셀프"].ToString().Equals(i["셀프"].ToString()) &&
                               row["시스템번호"].ToString().Equals(i["시스템번호"].ToString()))
                            {
                                i["RPA처리결과"] = regResult;
                                i["실패코드"] = regCode;
                                i["실패원인"] = regCause;
                                Console.WriteLine("PTN장치이력 업데이트 :: {0}, {1}, {2}", i["RPA처리결과"].ToString(), i["실패코드"].ToString(), i["실패원인"].ToString());
                                break;
                            }
                        }
                    }
                    catch (Exception E)
                    {
                        Console.WriteLine("<<E>>  PTNSystemReRegHist - 예외 발생 {0}\n{1}", E.GetType().ToString(), E.Message.ToString());
                    }
                }
            }

            /*
             * PTN 장치등록 NEOSS 지연 등록일경우 시스템 재등록하기 위함 
             * 등록 정보와 이력 정보를 비교해 재등록 대상을 발췌함
             */
            public class PTNSystemReReg : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<DataTable> regDt { get; set; }
                [Category("Input"), RequiredArgument]
                public InArgument<DataTable> regHistDt { get; set; }

                [Category("Output"), RequiredArgument]
                public OutArgument<DataTable> reRegDt { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> PTNSystemReReg - Start!");
                    reRegDt.Set(context, Work(regDt.Get(context), regHistDt.Get(context)));
                    Console.WriteLine("<<C>> PTNSystemReReg - End!");
                }

                private DataTable Work(DataTable reDt, DataTable regHistDt)
                {
                    try
                    {
                        DataTable resultDt = new DataTable();
                        resultDt.TableName = "장치등록정보";



                        foreach (DataColumn col in reDt.Columns)
                        {
                            resultDt.Columns.Add(col.ToString());
                        }

                        foreach (DataRow histRow in regHistDt.Rows)
                        {
                            if (histRow["실패코드"].ToString().Equals("R"))
                            {
                                foreach (DataRow regRow in reDt.Rows)
                                {
                                    if (histRow["장치설치위치"].ToString().Equals(regRow["장치설치위치"].ToString()) &&
                                        histRow["베이"].ToString().Equals(regRow["베이"].ToString()) &&
                                        histRow["셀프"].ToString().Equals(regRow["셀프"].ToString()) &&
                                        histRow["시스템번호"].ToString().Equals(regRow["시스템번호"].ToString()))
                                    {
                                        resultDt.ImportRow(regRow);
                                        break;
                                    }
                                }
                            }

                        }


                        return resultDt;
                    }
                    catch (Exception E)
                    {
                        Console.WriteLine("<<E>> PTNSystemReReg - 예외 발생 {0}\n{1}", E.GetType().ToString(), E.Message.ToString());
                        return null;
                    }
                }

            }



            /*
             * 엑셀 저장시 bay 0101 -> 101 정수 형태로 저장 되어 텍스트가 손실되는 셀을 올바르게 고치기 위해 사용
             *        bay       101 -> 0101
             *        system    1   -> 01
             *        slotnum   2   -> 02
             * 단, 원래 숫자형 셀에 대해서는 변환하지 않음 NumbericList에서 관리 
             */
            public class ConvertNumbericToString : CodeActivity
            {
                [Category("Input")]
                [Description("파라메터로 넘겨주는 데이터테이블은 결과 테이블의 컬럼을 정의할때 파라메터 테이블의 첫행을 기준으로 설정하기 때문에 이 점을 유의하고 사용하십시오")]
                [RequiredArgument]
                public InArgument<DataTable> param { get; set; }

                [Category("Option")]
                [DefaultValue(false)]
                public bool Addheader { get; set; }


                [Category("Output")]
                public OutArgument<DataTable> resultDT { get; set; }


                protected override void Execute(CodeActivityContext context)
                {
                    List<String> NumbericLIst = new List<String>();
                    NumbericLIst.Add("포트갯수");
                    NumbericLIst.Add("개수");
                    NumbericLIst.Add("시작타임슬롯");
                    DataTable argTable = param.Get(context);
                    DataTable t_resultDT = new DataTable();

                    t_resultDT = argTable;

                    // 새로운 태이블로 정의
                    DataTable RPA_DT = new DataTable();

                    // 인자로 넘어온 테이블은 헤더 설정이 없기때문에 첫번재 행을 헤더로 사용하기 위해 첫번째 행 저장
                    DataRow FirstRow = t_resultDT.Rows[0];

                    // 인자 테이블의 첫 번째 행을 새로 정의하는 테이블의 헤더로 설정하고, NULL값을 허용한다.
                    foreach (String column in FirstRow.ItemArray)
                    {
                        DataColumn RPA_Col = RPA_DT.Columns.Add(column, typeof(String));
                        RPA_Col.AllowDBNull = true;
                    }

                    // 헤더 설정을 마치고 첫 번째 행은 삭제하고 이후, 행부터는 새로 정의한 테이블에 add
                    t_resultDT.Rows.Remove(FirstRow);

                    // importRow시 기존 인자로 받은 테이블은 디폴트 테이블 헤더 설정을 갖고 있기 때문에 새로 정의한 테이블 헤더와 다르기 때문에 저장되지 않는다.
                    // 또, 같은 DataRow는 한 테이블에서만 가지고 있을 수 있기 때문에 서로 다른 테이블에서 같은 Row를 저장할 수 없다.
                    // -> 새로 DataRow를 생성해서 수작업으로 row를 세팅후 새로 정의한 테이블에 새 행으로 저장한다.
                    foreach (DataRow row in t_resultDT.Rows)
                    {
                        DataRow currentRow = RPA_DT.NewRow();

                        foreach (var x in row.ItemArray.Select((value, index) => new { value, index }))
                        {
                            currentRow[x.index] = x.value;
                        }
                        RPA_DT.Rows.Add(currentRow);
                    }


                    foreach (DataRow row in RPA_DT.Rows)
                    {

                        foreach (var x in row.ItemArray.Select((value, index) => new { value, index }))
                        {
                            String currentColName = row.Table.Columns[x.index].ColumnName;

                            /* String이 숫자형이고 길이가 1 또는 3이라면 베이 or 슬롯 or 시스템번호 etc.. 텍스트 형태 값에서 정수형으로 저장되었기 때문에 알맞은 자리수를 맞추기 위해서 
                             * 해당 조건을 만족하는 셀이라면 앞에 0을 붙인다. 
                             * 단, 위 조건과 무관한 기타 다른 행들에 대해서는 NumbericList에 컬럼명을 등록하고 해당 컬럼에 해당은 셀은 생략한다*/
                            if (Int32.TryParse(x.value.ToString(), out int numValue)
                                && (x.value.ToString().Length == 1 || x.value.ToString().Length == 3)
                                && !NumbericLIst.Contains(currentColName))
                            {
                                row.SetField(x.index, "0" + x.value);
                            }
                            else { }
                        }
                    }

                    resultDT.Set(context, RPA_DT);
                }
            }

            /*
             * RPA MSPP 등록시 지역 본부별로 엑셀파일 다운 받기 때문에 하나의 파일로 병합이 필요함.
             * DaTaTable n개를 1개로 병합시킴
             */
            public class DT_Merge : CodeActivity
            {

                [Category("Input")]
                [RequiredArgument]
                public InArgument<DataSet> param { get; set; }

                [Category("Output")]
                public OutArgument<DataTable> completeDt { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    DataSet dataTables = param.Get(context);
                    DataTable mergeDT = new DataTable();

                    int indexOfTables = 0;
                    foreach (DataTable pItem in dataTables.Tables)
                    {
                        // 첫번쨰 테이블은 헤더까지 저장하기 위해 별도 처리
                        if (indexOfTables == 0)
                        {
                            mergeDT.Merge(pItem);
                        }
                        else
                        {
                            pItem.Rows.RemoveAt(0);
                            mergeDT.Merge(pItem);
                        }
                        indexOfTables++;
                    }

                    completeDt.Set(context, mergeDT);

                }


            }
            /* ###############################################################
                FM불일치 MSPP 4형 장치등록정보 엑셀파일폼 생성하는 기능 수행
               ############################################################### */
            [Designer(typeof(CreateMismatchExcelFileActivityDesigner))]
            public class CreateMismatchExcelFile : CodeActivity
            {
                [DisplayName("FilePath"), Category("Input")]
                [RequiredArgument]
                [Description("파일 경로를 입력해주세요")]
                public InArgument<String> FilePath { get; set; }

                [Category("Input")]
                [RequiredArgument]
                public InArgument<String> equipType { get; set; }

                [Category("Output")]
                [Description("파일이 정상적으로 생성됐는지 여부")]
                public OutArgument<bool> CreateYN { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    String path = FilePath.Get(context);
                    String equiptype = equipType.Get(context);

                    if (equiptype.Equals("MSPP"))
                    {
                        CreateYN.Set(context, CreateMSPPFile(path));
                    }
                    else if (equiptype.Equals("PTS"))
                    {
                        CreateYN.Set(context, CreatePTNFile(path));
                    }
                    else
                    {
                        Console.WriteLine("<<E>> CreateMismatchExcelFile - equipType이 MSPP or PTS이 아닙니다. 파일을 생성하지 않음");
                        CreateYN.Set(context, false);
                    }


                }

                public bool CreateMSPPFile(String path)
                {
                    try
                    {
                        excel.Application app = new excel.Application();
                        // excel Display off 
                        app.Visible = false; // 백그라운드로 처리
                        app.DisplayAlerts = false; // 엑셀 경고창 표시안함

                        excel.Workbook workbook = app.Workbooks.Add();

                        /* 장치등록정보 시트 생성 */
                        excel.Worksheet worksheet_Equip = workbook.Worksheets.Item[1];

                        worksheet_Equip.Name = "장치등록정보";
                        worksheet_Equip.Cells[1, 1] = "본부";
                        worksheet_Equip.Cells[1, 2] = "TID";
                        worksheet_Equip.Cells[1, 3] = "관리국소";
                        worksheet_Equip.Cells[1, 4] = "fm설치위치";
                        worksheet_Equip.Cells[1, 5] = "장치설치위치";
                        worksheet_Equip.Cells[1, 6] = "장치대분류";
                        worksheet_Equip.Cells[1, 7] = "장치소분류";
                        worksheet_Equip.Cells[1, 8] = "베이";
                        worksheet_Equip.Cells[1, 9] = "셀프";
                        worksheet_Equip.Cells[1, 10] = "시스템번호";
                        worksheet_Equip.Cells[1, 11] = "서비스망";
                        worksheet_Equip.Cells[1, 12] = "사용용도";
                        worksheet_Equip.Cells[1, 13] = "자산조직";
                        worksheet_Equip.Cells[1, 14] = "제작사";
                        worksheet_Equip.Cells[1, 15] = "모델명";
                        worksheet_Equip.Cells[1, 16] = "KT자산여부";
                        worksheet_Equip.Cells[1, 17] = "망구분";
                        worksheet_Equip.Cells[1, 18] = "국사내여부";
                        worksheet_Equip.Cells[1, 19] = "설치위치변경여부";

                        // 컬럼 가운데 정렬하고 배경색 지정
                        excel.Range HorizontalAlignmentCell_Equip = worksheet_Equip.Range["A1", "S1"];
                        HorizontalAlignmentCell_Equip.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
                        HorizontalAlignmentCell_Equip.Interior.Color = excel.XlRgbColor.rgbYellow;

                        // 텍스트형 셀 설정 
                        excel.Range TextCell_Equip = worksheet_Equip.Range["A1", "J299"];
                        TextCell_Equip.NumberFormat = "@";


                        /* 유니트등록정보 시트 생성 */
                        excel.Worksheet worksheet_Unit = workbook.Worksheets.Add(After: worksheet_Equip);
                        worksheet_Unit.Name = "유니트등록정보";
                        worksheet_Unit.Cells[1, 1] = "설치위치";
                        worksheet_Unit.Cells[1, 2] = "장치명";
                        worksheet_Unit.Cells[1, 3] = "시스템번호";
                        worksheet_Unit.Cells[1, 4] = "슬롯범위";
                        worksheet_Unit.Cells[1, 5] = "유니트명";
                        worksheet_Unit.Cells[1, 6] = "유니트구분";
                        worksheet_Unit.Cells[1, 7] = "대역폭";
                        worksheet_Unit.Cells[1, 8] = "포트갯수";

                        // 컬럼 가운데 정렬하고 배경색 지정
                        excel.Range HorizontalAlignmentCell_Unit = worksheet_Unit.Range["A1", "H1"];
                        HorizontalAlignmentCell_Unit.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
                        HorizontalAlignmentCell_Unit.Interior.Color = excel.XlRgbColor.rgbYellow;

                        // 텍스트형 셀 설정 
                        excel.Range TextCell_Unit = worksheet_Unit.Range["A1", "H299"];
                        TextCell_Unit.NumberFormat = "@";


                        /* 캐리어등록정보 시트 생성 */
                        excel.Worksheet worksheet_Carrier = workbook.Worksheets.Add(After: worksheet_Unit);
                        worksheet_Carrier.Name = "캐리어등록정보";
                        worksheet_Carrier.Cells[1, 1] = "하위설치위치";
                        worksheet_Carrier.Cells[1, 2] = "장치명";
                        worksheet_Carrier.Cells[1, 3] = "시스템";
                        worksheet_Carrier.Cells[1, 4] = "하위포트명";
                        worksheet_Carrier.Cells[1, 5] = "상위설치위치";
                        worksheet_Carrier.Cells[1, 6] = "상위장치소분류";
                        worksheet_Carrier.Cells[1, 7] = "상위장치명";
                        worksheet_Carrier.Cells[1, 8] = "상위포트명";
                        worksheet_Carrier.Cells[1, 9] = "캐리어번호";
                        worksheet_Carrier.Cells[1, 10] = "캐리어구분";


                        // 컬럼 가운데 정렬하고 배경색 지정
                        excel.Range HorizontalAlignmentCell_Carrier = worksheet_Carrier.Range["A1", "J1"];
                        HorizontalAlignmentCell_Carrier.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
                        HorizontalAlignmentCell_Carrier.Interior.Color = excel.XlRgbColor.rgbYellow;

                        // 텍스트형 셀 설정 
                        excel.Range TextCell_Carrier = worksheet_Carrier.Range["A1", "J299"];
                        TextCell_Carrier.NumberFormat = "@";

                        /* 전송로등록정보 시트 생성 */
                        excel.Worksheet worksheet_Transline = workbook.Worksheets.Add(After: worksheet_Carrier);
                        worksheet_Transline.Name = "전송로등록정보";
                        worksheet_Transline.Cells[1, 1] = "하위설치위치";
                        worksheet_Transline.Cells[1, 2] = "장치명";
                        worksheet_Transline.Cells[1, 3] = "시스템";
                        worksheet_Transline.Cells[1, 4] = "하위포트명";
                        worksheet_Transline.Cells[1, 5] = "상위설치위치";
                        worksheet_Transline.Cells[1, 6] = "캐리어번호";
                        worksheet_Transline.Cells[1, 7] = "계위";
                        worksheet_Transline.Cells[1, 8] = "시작타임슬롯";
                        worksheet_Transline.Cells[1, 9] = "개수";
                        worksheet_Transline.Cells[1, 10] = "전용회선번호";
                        worksheet_Transline.Cells[1, 11] = "Drop연결";


                        // 컬럼 가운데 정렬하고 배경색 지정
                        excel.Range HorizontalAlignmentCell_Transline = worksheet_Transline.Range["A1", "K1"];
                        HorizontalAlignmentCell_Transline.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
                        HorizontalAlignmentCell_Transline.Interior.Color = excel.XlRgbColor.rgbYellow;

                        // 텍스트형 셀 설정 
                        excel.Range TextCell_Transline = worksheet_Transline.Range["A1", "K299"];
                        TextCell_Transline.NumberFormat = "@";


                        if (File.Exists(@path))
                        {
                            File.Delete(path);
                            Console.WriteLine(path + " 파일이 이미 존해해서 덮어씌움");
                        }



                        workbook.SaveAs(Filename: @path);
                        workbook.Close();
                        app.Quit();
                    }
                    catch (Exception exception)
                    {

                        Console.WriteLine(exception.Message);
                        return false;
                    }



                    return true;
                }

                public bool CreatePTNFile(String path)
                {
                    try
                    {
                        excel.Application app = new excel.Application();
                        // excel Display off 
                        app.Visible = false; // 백그라운드로 처리
                        app.DisplayAlerts = false; // 엑셀 경고창 표시안함

                        excel.Workbook workbook = app.Workbooks.Add();

                        /* 장치등록정보 시트 생성 */
                        excel.Worksheet worksheet_Equip = workbook.Worksheets.Item[1];

                        worksheet_Equip.Name = "장치등록정보";
                        worksheet_Equip.Cells[1, 1] = "본부";
                        worksheet_Equip.Cells[1, 2] = "PTN노드명(TID)";
                        worksheet_Equip.Cells[1, 3] = "관리국소";
                        worksheet_Equip.Cells[1, 4] = "fm설치위치";
                        worksheet_Equip.Cells[1, 5] = "장치설치위치";
                        worksheet_Equip.Cells[1, 6] = "장치대분류";
                        worksheet_Equip.Cells[1, 7] = "장치소분류";
                        worksheet_Equip.Cells[1, 8] = "베이";
                        worksheet_Equip.Cells[1, 9] = "셀프";
                        worksheet_Equip.Cells[1, 10] = "시스템번호";
                        worksheet_Equip.Cells[1, 11] = "서비스망";
                        worksheet_Equip.Cells[1, 12] = "사용용도";
                        worksheet_Equip.Cells[1, 13] = "자산조직";
                        worksheet_Equip.Cells[1, 14] = "제작사";
                        worksheet_Equip.Cells[1, 15] = "모델명";
                        worksheet_Equip.Cells[1, 16] = "KT자산여부";
                        worksheet_Equip.Cells[1, 17] = "국사내여부";
                        worksheet_Equip.Cells[1, 18] = "설치위치변경여부";

                        // 컬럼 가운데 정렬하고 배경색 지정
                        excel.Range HorizontalAlignmentCell_Equip = worksheet_Equip.Range["A1", "S1"];
                        HorizontalAlignmentCell_Equip.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
                        HorizontalAlignmentCell_Equip.Interior.Color = excel.XlRgbColor.rgbYellow;

                        // 텍스트형 셀 설정 
                        excel.Range TextCell_Equip = worksheet_Equip.Range["A1", "J299"];
                        TextCell_Equip.NumberFormat = "@";


                        if (File.Exists(@path))
                        {
                            File.Delete(path);
                            Console.WriteLine(path + " 파일이 이미 존해해서 덮어씌움");
                        }



                        workbook.SaveAs(Filename: @path);
                        workbook.Close();
                        app.Quit();
                    }
                    catch (Exception exception)
                    {

                        Console.WriteLine(exception.Message);
                        return false;
                    }



                    return true;
                }
            }

        }// [FmMismatch] end of excel

        namespace Outlook
        {
            public static class OutlookManager
            {

                public static outlook.Application outlookApp = new outlook.Application();
                public static outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");


                public static bool ReadAllMail(String id, String passwd, String mailFolderName)
                {
                    if (mailFolderName == null)
                    {
                        Console.WriteLine("메일함 이름이 Null 입니다.");
                        return false;
                    }

                    outlookNamespace.Logon("Outlook", passwd, false, false);
                    //System.Threading.Thread.Sleep(3000);
                    //outlookNamespace.SendAndReceive(false);
                    outlook.MAPIFolder myInbox;// = outlookNamespace.GetDefaultFolder(outlook.OlDefaultFolders.olFolderInbox);

                    myInbox = outlookNamespace.Folders[id].Folders[mailFolderName];
                    //Console.WriteLine("folders : {0}", myInbox.Folders.Count);
                    //Console.WriteLine("Account : {0}", outlookNamespace.Accounts.Count);
                    //Console.WriteLine("Name : {0}", outlookNamespace.Accounts[1].DisplayName);


                    //outlook.MailItem NewMail = (outlook.MailItem)outlookApp.CreateItem(outlook.OlItemType.olMailItem);

                    try
                    {
                        foreach (object item in myInbox.Items)
                        {

                            outlook.MailItem newEmail = item as outlook.MailItem;

                            if (newEmail != null)
                            {
                                newEmail.UnRead = false;
                                newEmail.Save();
                            }
                        }


                        return true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message.ToString());
                        outlookApp.Quit();
                        return false;
                    }
                }
                public static void RefleshMailFolder(String id, String passwd)
                {
                    try
                    {
                        outlookNamespace.Logon("Outlook", passwd, false, false);
                        outlookNamespace.SendAndReceive(false);
                        Console.WriteLine("[{0}] RefleshOutlook - SendAndReceive 를 수행했습니다.", DateTime.Now.ToString("hh:mm:ss"));


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message.ToString());
                        outlookApp.Quit();
                    }

                }
            }

            [Designer(typeof(ReadAllMailFolderActivityDesigner))]
            public class ReadAllMailFolder : CodeActivity
            {
                [DisplayName("ID"), Category("Input")]
                public InArgument<String> ID { get; set; }

                [PasswordPropertyText(true)]
                [DataType(DataType.Password)]
                [DisplayName("Passwd"), Category("Input")]
                [RequiredArgument]
                public InArgument<String> Passwd { get; set; }

                [DisplayName("MailFolderName"), Category("Input")]
                [RequiredArgument]
                public InArgument<String> MailFolderName { get; set; }

                [Category("Output")]
                public OutArgument<bool> ResultYN { get; set; }
                protected override void Execute(CodeActivityContext context)
                {

                    Console.WriteLine("<<C>> ReadAllMailFolder Start!");
                    String id = ID.Get(context);
                    String passwd = Passwd.Get(context);
                    String foldername = MailFolderName.Get(context);

                    ResultYN.Set(context, OutlookManager.ReadAllMail(id, passwd, foldername));
                    Console.WriteLine("<<C>> ReadAllMailFolder End");
                }
            }

            /*
             * RPA - UIPATH outlook 관련 Activity 사용시 새 메일에 관해서 새로고침 기능이 없어서 만듬
             * 현재 메일함 새로고침 시켜줌 'Outlook 리본탭에 모든 폴더 보내기/받기' 기능에 해당함
             */
            [Designer(typeof(RefleshOutlookActivityDesigner))]
            public class RefleshOutlook : CodeActivity
            {
                [DisplayName("ID"), Category("Input")]
                [Description("Outlook 새 메일함을 업데이트 해주는 기능입니다. ")]
                [RequiredArgument]
                public InArgument<String> ID { get; set; }

                [PasswordPropertyText(true)]
                [DataType(DataType.Password)]
                [DisplayName("PASSWD"), Category("Input")]
                [RequiredArgument]
                public InArgument<String> PASSWD { get; set; }

                protected override void Execute(CodeActivityContext context)
                {

                    Console.WriteLine("<<C>> RefleshOutlook Start!");
                    String id = ID.Get(context);
                    String passwd = PASSWD.Get(context);

                    if (id == null || passwd == null)
                    {
                        Console.WriteLine("RefleshOutlook - 입력 파라메터가 null 입니다.");
                    }
                    else
                    {
                        OutlookManager.RefleshMailFolder(id, passwd);
                    }
                    Console.WriteLine("<<C>> RefleshOutlook End");

                }


            }
        } // [FmMismatch] end of Outlook

        namespace Converter
        {

            /*
             * JsonString을 DataSet으로 변환 FM 불일치에 한해서 동작 주의!!
             */
            public class RPARegInfo2DataSet : CodeActivity
            {
                [Category("Input")]
                public InArgument<String> string4Jsontype { get; set; }

                [Category("Output")]
                public OutArgument<DataSet> resultDataSet { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    resultDataSet.Set(context, JsonConverter(string4Jsontype.Get(context)));

                }

                private DataSet JsonConverter(String targetStr)
                {
                    if (targetStr == null)
                    {
                        Console.WriteLine("대상 문자열이 null 입니다.");
                    }
                    /*
                     * 일반적인 Json 형태의 문자열을 받아오지 않기 때문에 아래 별도 작업이 필요하다.
                     * json구조 자체를 String으로 변환후 다시 K/V 형태의 오브젝트에 값을 할당후 또, 다시 
                     * String 형태로 변환이 되기 때문.
                     * ---------- Before ----------
                     * {
                     *      "name" : "junyoung",
                     *      "age"  : "28"
                     * }
                     * ---------- After -----------
                     * "{
                     *      \"name\" : \"junyoung\",
                     *      \"age\"  : \"28\"
                     * }"
                     * after후, 쌍따옴표에 백슬래쉬가 붙여짐.
                     */

                    // STEP 0 : "{ 을 { 로, }"을 }로 바꾼다. 
                    targetStr = targetStr.Replace("\"{", "{");
                    targetStr = targetStr.Replace("}\"", "}");
                    Console.WriteLine("step1 : {0}", targetStr);

                    // STEP 1 : '[ 을 [ 로, ]'을 ]로 바꾼다.
                    targetStr = targetStr.Replace("'[", "[");
                    targetStr = targetStr.Replace("]'", "]");
                    Console.WriteLine("step2 : {0}", targetStr);

                    // STEP 2 : '을 \"로 바꾼다.
                    targetStr = targetStr.Replace("'", "\"");
                    Console.WriteLine("step3 : {0}", targetStr);

                    // STEP 3 : \n을 지운다.
                    targetStr = targetStr.Replace("\n", "");
                    Console.WriteLine("step4 : {0}", targetStr);

                    // STEP 4 : 공백을 지운다.
                    //targetStr = targetStr.Replace(" ", "");
                    //Console.WriteLine("step5 : {0}", targetStr);

                    try
                    {
                        // JSON 확인
                        Js.JObject targetJobj = Js.JObject.Parse(targetStr);
                        Console.WriteLine("{0}", targetJobj.ToString());

                        Js.JObject jroot = (Js.JObject)targetJobj["output"];
                        Js.JObject jdata = (Js.JObject)jroot["resultMsg"];

                        String equipType = ((String)jroot["result"]).Equals("0") ? "MSPP" : ((String)jroot["result"]).Equals("1") ? "PTS" : ((String)jroot["result"]);

                        if (equipType.Equals("MSPP"))
                        {
                            DataSet Dtset = new DataSet();

                            /* 0. 장치등록정보 테이블 생성후 컬럼추가 */

                            Js.JArray equipJA = (Js.JArray)jdata["장치등록정보"];
                            DataTable equipDT = new DataTable();
                            equipDT.TableName = "장치등록정보";
                            foreach (String ecitem in GlobalConstants.RPAEQUIPCOLUMN)
                            {
                                DataColumn equipcol = new DataColumn();
                                equipcol.DataType = System.Type.GetType("System.String");
                                equipcol.ColumnName = ecitem;

                                equipDT.Columns.Add(equipcol);
                            }
                            Console.WriteLine("equip Column \n{0}", equipDT.Columns.ToString());

                            foreach (var Equipitem in equipJA)
                            {
                                DataRow equipRow = equipDT.NewRow();
                                equipRow["본부"] = Equipitem["region"];
                                equipRow["TID"] = Equipitem["tid"];
                                equipRow["관리국소"] = Equipitem["officename"];
                                equipRow["fm설치위치"] = Equipitem["fminstlocation"];
                                equipRow["장치설치위치"] = Equipitem["instlocation"];
                                equipRow["장치대분류"] = Equipitem["mainclscode"];
                                equipRow["장치소분류"] = Equipitem["subclscode"];
                                equipRow["베이"] = Equipitem["bay"];
                                equipRow["셀프"] = Equipitem["shelf"];
                                equipRow["시스템번호"] = Equipitem["system"];
                                equipRow["서비스망"] = Equipitem["svcnetcode"];
                                equipRow["사용용도"] = Equipitem["usagecode"];
                                equipRow["자산조직"] = Equipitem["assetorg"];
                                equipRow["제작사"] = Equipitem["vendor"];
                                equipRow["모델명"] = Equipitem["model"];
                                equipRow["KT자산여부"] = Equipitem["isktasset"];
                                equipRow["망구분"] = Equipitem["netcode"];
                                equipRow["국사내여부"] = Equipitem["insideoffice"];
                                equipRow["설치위치변경여부"] = "";

                                equipDT.Rows.Add(equipRow);
                            }
                            Console.WriteLine("equip DataTable \n{0}", equipDT.Rows.ToString());

                            /* 1. 유니트등록정보 테이블 생성후 컬럼추가 */
                            Js.JArray unitJA = (Js.JArray)jdata["유니트등록정보"];
                            DataTable unitDT = new DataTable();
                            unitDT.TableName = "유니트등록정보";
                            foreach (String ucitem in GlobalConstants.RPAUNITCOLUMN)
                            {
                                DataColumn unitcol = new DataColumn();
                                unitcol.DataType = System.Type.GetType("System.String");
                                unitcol.ColumnName = ucitem;

                                unitDT.Columns.Add(unitcol);
                            }
                            Console.WriteLine("unit Column \n{0}", unitDT.Columns.ToString());

                            foreach (var Unititem in unitJA)
                            {
                                DataRow unitRow = unitDT.NewRow();
                                unitRow["설치위치"] = Unititem["site"];
                                unitRow["장치명"] = Unititem["sysname"];
                                unitRow["시스템번호"] = Unititem["system"];
                                unitRow["슬롯범위"] = Unititem["slotnum"];
                                unitRow["유니트명"] = Unititem["unitmodel"];
                                unitRow["유니트구분"] = Unititem["unittype"];
                                unitRow["대역폭"] = Unititem["bandwidth"];
                                unitRow["포트갯수"] = Unititem["portcount"];

                                unitDT.Rows.Add(unitRow);
                            }
                            Console.WriteLine("unit DataTable \n{0}", unitDT.Rows.ToString());

                            /* 2. 캐리어등록정보 테이블 생성후 컬럼추가 */
                            Js.JArray carrierJA = (Js.JArray)jdata["캐리어등록정보"];
                            DataTable carrierDT = new DataTable();
                            carrierDT.TableName = "캐리어등록정보";
                            foreach (String ccitem in GlobalConstants.RPACARRIERCOLUMN)
                            {
                                DataColumn carriercol = new DataColumn();
                                carriercol.DataType = System.Type.GetType("System.String");
                                carriercol.ColumnName = ccitem;

                                carrierDT.Columns.Add(carriercol);
                            }
                            Console.WriteLine("carrier Column \n{0}", carrierDT.Columns.ToString());

                            foreach (var carrieritem in carrierJA)
                            {
                                DataRow carrierRow = carrierDT.NewRow();
                                carrierRow["하위설치위치"] = carrieritem["lowersite"];
                                carrierRow["장치명"] = carrieritem["lowersysname"];
                                carrierRow["시스템"] = carrieritem["lowersystem"];
                                carrierRow["하위포트명"] = carrieritem["lowerfmrssup"];
                                carrierRow["상위설치위치"] = carrieritem["uppersite"];
                                // 2020.07.07 상위 장치 소분류에 공백 들어가 있는 경우 있어서 오류남.. 
                                carrierRow["상위장치소분류"] = carrieritem["uppersubclscode"].ToString().Replace(" ","");
                                carrierRow["상위장치명"] = carrieritem["uppersysname"];
                                carrierRow["상위포트명"] = carrieritem["upperfmrssup"];
                                carrierRow["캐리어번호"] = carrieritem["carriernum"];
                                carrierRow["캐리어구분"] = carrieritem["carriertype"];

                                carrierDT.Rows.Add(carrierRow);
                            }
                            Console.WriteLine("carrier DataTable \n{0}", carrierDT.Rows.ToString());

                            /* 3. 전송로등록정보 테이블 생성후 컬럼추가 */
                            Js.JArray translineJA = (Js.JArray)jdata["전송로등록정보"];
                            DataTable translineDT = new DataTable();
                            translineDT.TableName = "전송로등록정보";
                            foreach (String tcitem in GlobalConstants.RPATRANSLINECOLUMN)
                            {
                                DataColumn translinecol = new DataColumn();
                                translinecol.DataType = System.Type.GetType("System.String");
                                translinecol.ColumnName = tcitem;

                                translineDT.Columns.Add(translinecol);
                            }
                            Console.WriteLine("transline Column \n{0}", translineDT.Columns.ToString());

                            foreach (var translineitem in translineJA)
                            {
                                DataRow translineRow = translineDT.NewRow();
                                translineRow["하위설치위치"] = translineitem["lowersite"];
                                translineRow["장치명"] = translineitem["lowersysname"];
                                translineRow["시스템"] = translineitem["lowersystem"];
                                translineRow["하위포트명"] = translineitem["lowerfmrssup"];
                                translineRow["상위설치위치"] = translineitem["uppersite"];
                                translineRow["캐리어번호"] = translineitem["carriernum"];
                                translineRow["계위"] = translineitem["layertype"];
                                translineRow["시작타임슬롯"] = translineitem["starttimeslot"];
                                translineRow["개수"] = translineitem["timeslotcount"];
                                translineRow["전용회선번호"] = translineitem["leasedline"];
                                translineRow["Drop연결"] = translineitem["dropfmrssup"];

                                translineDT.Rows.Add(translineRow);
                            }
                            Console.WriteLine("transline DataTable \n{0}", translineDT.Rows.ToString());

                            Dtset.Tables.Add(equipDT);
                            Dtset.Tables.Add(unitDT);
                            Dtset.Tables.Add(carrierDT);
                            Dtset.Tables.Add(translineDT);

                            return Dtset;
                        }
                        else if (equipType.Equals("PTS"))
                        {
                            DataSet Dtset = new DataSet();

                            /* 0. 장치등록정보 테이블 생성후 컬럼추가 */

                            Js.JArray equipJA = (Js.JArray)jdata["장치등록정보"];
                            DataTable equipDT = new DataTable();
                            equipDT.TableName = "장치등록정보";
                            foreach (String ecitem in GlobalConstants.RPAPTNEQUIPCOLUMN)
                            {
                                DataColumn equipcol = new DataColumn();
                                equipcol.DataType = System.Type.GetType("System.String");
                                equipcol.ColumnName = ecitem;

                                equipDT.Columns.Add(equipcol);
                            }
                            Console.WriteLine("equip Column \n{0}", equipDT.Columns.ToString());

                            foreach (var Equipitem in equipJA)
                            {
                                DataRow equipRow = equipDT.NewRow();
                                equipRow["본부"] = "";
                                equipRow["PTN노드명(TID)"] = Equipitem["tid"];
                                equipRow["관리국소"] = Equipitem["officename"];
                                equipRow["fm설치위치"] = Equipitem["fminstlocation"];
                                equipRow["장치설치위치"] = Equipitem["instlocation"];
                                equipRow["장치대분류"] = Equipitem["mainclscode"];
                                equipRow["장치소분류"] = Equipitem["subclscode"];
                                equipRow["베이"] = Equipitem["bay"];
                                equipRow["셀프"] = Equipitem["shelf"];
                                equipRow["시스템번호"] = Equipitem["system"];
                                equipRow["서비스망"] = Equipitem["svcnetcode"];
                                equipRow["사용용도"] = Equipitem["usagecode"];
                                equipRow["자산조직"] = Equipitem["assetorg"];
                                equipRow["제작사"] = Equipitem["vendor"];
                                equipRow["모델명"] = Equipitem["model"];
                                equipRow["KT자산여부"] = Equipitem["isktasset"];
                                equipRow["국사내여부"] = Equipitem["insideoffice"];
                                equipRow["설치위치변경여부"] = "";

                                equipDT.Rows.Add(equipRow);
                            }
                            Console.WriteLine("equip DataTable \n{0}", equipDT.Rows.ToString());

                            Dtset.Tables.Add(equipDT);
                            return Dtset;
                        }
                        else
                        {
                            Console.WriteLine("[ResultCode : {0}] - 결과 데이터를 테이블로 생성할 수 없습니다,", equipType);
                            return null;
                        }



                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        return null;
                    }

                }

            }

            /*
             * FM 불일치 장치등록 RPA 처리 이력 데이터를 RPC요청 하기위한 JSON Type으로 변환
             */
            public class RPAHistory2Json : CodeActivity
            {
                [Category("Input")]
                [RequiredArgument]
                public InArgument<DataSet> inputDataSet { get; set; }

                [Category("input")]
                [RequiredArgument]
                public InArgument<String> equipType { get; set; }

                [Category("Output")]
                public OutArgument<String> resultString4jsontype { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    String equiptype = equipType.Get(context);

                    if (equiptype.Equals("MSPP"))
                    {
                        resultString4jsontype.Set(context, MSPPDataSet2JsonConverter(inputDataSet.Get(context)));
                    }
                    else if (equiptype.Equals("PTS"))
                    {
                        resultString4jsontype.Set(context, PTNDataSet2JsonConverter(inputDataSet.Get(context)));
                    }
                    else
                    {
                        Console.WriteLine("<<E>> RPAHistory2Json - equipType이 MSPP or PTS이 아닙니다.");
                    }

                }

                private String MSPPDataSet2JsonConverter(DataSet pinputDataSet)
                {
                    try
                    {
                        // 0. 장치 등록 이력 데이터 테이블 K,V 형태로 변환
                        DataTable equipHistoryDT = pinputDataSet.Tables["장치등록이력"];

                        List<Dictionary<String, String>> equipDataList = new List<Dictionary<String, String>>();

                        foreach (DataRow equipRowitem in equipHistoryDT.Rows)
                        {
                            Dictionary<String, String> equiptmpdic = new Dictionary<string, string>();
                            foreach (DataColumn colum in equipHistoryDT.Columns)
                            {
                                equiptmpdic.Add(colum.ToString(), equipRowitem[colum].ToString());
                            }
                            equipDataList.Add(equiptmpdic);
                        }

                        // 1. 유니트 등록 이력 데이터 테이블 K,V 형태로 변환
                        DataTable unitHistoryDT = pinputDataSet.Tables["유니트등록이력"];

                        List<Dictionary<String, String>> unitDataList = new List<Dictionary<String, String>>();

                        foreach (DataRow unitRowitem in unitHistoryDT.Rows)
                        {
                            Dictionary<String, String> unittmpdic = new Dictionary<string, string>();
                            foreach (DataColumn colum in unitHistoryDT.Columns)
                            {
                                unittmpdic.Add(colum.ToString(), unitRowitem[colum].ToString());
                            }
                            unitDataList.Add(unittmpdic);
                        }

                        // 2. 캐리어 등록 이력 데이터 테이블 K,V 형태로 변환
                        DataTable carrierHistoryDT = pinputDataSet.Tables["캐리어등록이력"];

                        List<Dictionary<String, String>> carrierDataList = new List<Dictionary<String, String>>();

                        foreach (DataRow carrierRowitem in carrierHistoryDT.Rows)
                        {
                            Dictionary<String, String> carriertmpdic = new Dictionary<string, string>();
                            foreach (DataColumn colum in carrierHistoryDT.Columns)
                            {
                                carriertmpdic.Add(colum.ToString(), carrierRowitem[colum].ToString());
                            }
                            carrierDataList.Add(carriertmpdic);
                        }

                        // 3. 캐리어 등록 이력 데이터 테이블 K,V 형태로 변환
                        DataTable translineHistoryDT = pinputDataSet.Tables["전송로등록이력"];

                        List<Dictionary<String, String>> translineDataList = new List<Dictionary<String, String>>();

                        foreach (DataRow translineRowitem in translineHistoryDT.Rows)
                        {
                            Dictionary<String, String> translinetmpdic = new Dictionary<string, string>();
                            foreach (DataColumn colum in translineHistoryDT.Columns)
                            {
                                translinetmpdic.Add(colum.ToString(), translineRowitem[colum].ToString());
                            }
                            translineDataList.Add(translinetmpdic);
                        }


                        // 이력 데이터 json string 형태로 만들어서 리턴

                        Dictionary<String, object> tmpHistoryObj = new Dictionary<string, object>();

                        tmpHistoryObj.Add("장치등록이력", equipDataList);
                        tmpHistoryObj.Add("유니트등록이력", unitDataList);
                        tmpHistoryObj.Add("캐리어등록이력", carrierDataList);
                        tmpHistoryObj.Add("전송로등록이력", translineDataList);

                        String rpaHist_dictionary2JsonStr = JsonConvert.SerializeObject(tmpHistoryObj, Formatting.Indented);
                        Console.WriteLine(rpaHist_dictionary2JsonStr);
                        String changeDoubleQouteToSingleQoute = rpaHist_dictionary2JsonStr.Replace("\"", "'");
                        String removeLF = changeDoubleQouteToSingleQoute.Replace("\n", "");
                        String removeCR = removeLF.Replace("\r", "");
                        //String removeSpace = removeCR.Replace(" ", "");

                        Console.WriteLine(removeCR);

                        return removeCR;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        return null;
                    }

                }
                private String PTNDataSet2JsonConverter(DataSet pinputDataSet)
                {
                    try
                    {
                        // 0. 장치 등록 이력 데이터 테이블 K,V 형태로 변환
                        DataTable equipHistoryDT = pinputDataSet.Tables["장치등록이력"];

                        List<Dictionary<String, String>> equipDataList = new List<Dictionary<String, String>>();

                        foreach (DataRow equipRowitem in equipHistoryDT.Rows)
                        {
                            Dictionary<String, String> equiptmpdic = new Dictionary<string, string>();
                            foreach (DataColumn colum in equipHistoryDT.Columns)
                            {
                                equiptmpdic.Add(colum.ToString(), equipRowitem[colum].ToString());
                            }
                            equipDataList.Add(equiptmpdic);
                        }

                        // 이력 데이터 json string 형태로 만들어서 리턴

                        Dictionary<String, object> tmpHistoryObj = new Dictionary<string, object>();

                        tmpHistoryObj.Add("장치등록이력", equipDataList);

                        String rpaHist_dictionary2JsonStr = JsonConvert.SerializeObject(tmpHistoryObj, Formatting.Indented);
                        Console.WriteLine(rpaHist_dictionary2JsonStr);
                        String changeDoubleQouteToSingleQoute = rpaHist_dictionary2JsonStr.Replace("\"", "'");
                        String removeLF = changeDoubleQouteToSingleQoute.Replace("\n", "");
                        String removeCR = removeLF.Replace("\r", "");
                        //String removeSpace = removeCR.Replace(" ", "");

                        Console.WriteLine(removeCR);

                        return removeCR;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        return null;
                    }

                }
            }

            /* 
            ###############################################################
            UIPATH에서 메타문자 앞에 '&amp'처럼 자동으로 붙여서 생성됨
            특히, 자산조직 선택시 &붙여진 하위 자산조직 선택하는 과정에서
            셀렉터에서 &를 인식하지 못하고 &amp가 붙여저서 셀렉트하지 못하는 문제발생 
            때문에 &amp를 다시 &로 변경해주는 후처리 필요
            (UIPATH에서 강제로 붙이기 때문에 옵션 ON/OFF 으로 처리가 불가)
           ############################################################### */
            public class MetaStringConverter : CodeActivity
            {
                enum meta { amp, lt, gt, quot, sharp, apostrophe };
                [Category("Input")]
                [RequiredArgument]
                public InArgument<String> param { get; set; }

                [Category("Output")]
                public OutArgument<String> outData { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    String target = param.Get(context);

                    foreach (int num in Enum.GetValues(typeof(meta)))
                    {
                        switch (num)
                        {
                            case 0:
                                if (target.Contains("&"))
                                    target = target.Replace("&", "&amp;");
                                break;
                            case 1:
                                if (target.Contains("<"))
                                    target = target.Replace("<", "&lt;");
                                break;
                            case 2:
                                if (target.Contains(">"))
                                    target = target.Replace(">", "&gt;");
                                break;
                            case 3:
                                if (target.Contains("\""))
                                    target = target.Replace("\"", "&quot;");
                                break;
                            case 4:
                                if (target.Contains("#"))
                                    target = target.Replace("#", "&#035;");
                                break;
                            case 5:
                                if (target.Contains("'"))
                                    target = target.Replace("'", "&#039;");
                                break;

                        }
                    }

                    outData.Set(context, target);

                }
            }

            /* ###############################################################
                DataTable to JSON Format String Converter
               ###############################################################*/
            public class DataTableToJSON : CodeActivity
            {
                [Category("Input")]
                public InArgument<DataTable> datatable { get; set; }

                // 임시 테스트용 실제 기능에서는 UIPath에 결과 리턴 없을수도
                [Category("Output")]
                public OutArgument<String> resultStr { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    //subUtil util = new subUtil();
                    DataTable dt = datatable.Get(context);

                    List<Dictionary<String, String>> list = new List<Dictionary<String, String>>();
                    //Dictionary<String, String> 

                    foreach (DataRow row in dt.Rows)
                    {
                        Dictionary<String, String> tmp = new Dictionary<string, string>();
                        foreach (DataColumn colum in dt.Columns)
                        {
                            tmp.Add(colum.ToString(), row[colum].ToString());
                        }
                        list.Add(tmp);
                    }

                    foreach (Dictionary<String, String> dic in list)
                    {
                        Console.WriteLine("{0}", dic["본부"]);
                    }
                    String resultJson = JsonConvert.SerializeObject(list, Formatting.Indented);

                    Console.WriteLine(resultJson);

                    resultStr.Set(context, resultJson);


                }
            }
        } // [FmMismatch] end of Converter

        namespace etc
        {
            /* ###############################################################
            FM불일치 4형장치등록 프로세서 실행 요일 관리
           ###############################################################*/
            public class RPA_ExcuteYN : CodeActivity
            {

                [Category("Input")]
                public InArgument<DateTime> today { get; set; }

                [Category("Output")]
                public OutArgument<bool> resultYN { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    DateTime _today = today.Get(context);
                    bool _result;

                    switch (_today.DayOfWeek)
                    {
                        case DayOfWeek.Monday:
                            _result = true;
                            break;
                        case DayOfWeek.Tuesday:
                            _result = true;
                            break;
                        case DayOfWeek.Wednesday:
                            _result = true;
                            break;
                        case DayOfWeek.Thursday:
                            _result = true;
                            break;
                        case DayOfWeek.Friday:
                            _result = true;
                            break;
                        default:
                            _result = false;
                            break;
                    }

                    resultYN.Set(context, _result);
                }

            }
        } // [FmMismatch] end of etc
    }

    namespace EMS
    {
        namespace RPC
        {
            public class EmsRpcRequest : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<String> URL { get; set; }

                [Category("Input")]
                public InArgument<String> Parameter { get; set; }

                [Category("Output")]
                public OutArgument<String> Response { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> EmsRpcRequest - EMS 등록정보 RPC 요청 Start!");
                    Response.Set(context, work(this.URL.Get(context), this.Parameter.Get(context)));
                    Console.WriteLine("<<C>> EmsRpcRequest - End!");

                }

                private String work(String URL, String parameter)
                {
                    String rpcResultBody = null;
                    try
                    {
                        Uri uri = new Uri(URL);
                        var httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);

                        httpWebRequest.ServicePoint.Expect100Continue = false;
                        httpWebRequest.Method = "POST";
                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Accept = "application/json";

                        Console.WriteLine(parameter);
                        using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            streamWriter.Write(parameter);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();
                            Console.WriteLine("RPC 요청 결과 = \n{0}", result.ToString());
                            rpcResultBody = result.ToString();
                        }

                    }
                    catch (Exception E)
                    {
                        Console.WriteLine("<<E>> EmsRpcRequest :: 예외 타입 - {0} / {1}", E.GetType(), E.Message);
                    }
                    finally
                    {

                    }


                    return rpcResultBody;
                }
            }
        }
        namespace Biz
        {

            /*
             * 텔레필드 EMS 등록정보중 NeName을 분리하여 입력받도록 되어 있기 때문에 파싱한다. 
             */
            public class TFSysnameParser : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<String> targetStr { get; set; }

                [Category("Output")]
                public OutArgument<Dictionary<String, String>> result { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    Dictionary<String, String> resultDic = new Dictionary<string, string>();
                    //Regex regex = new Regex(@"[\w|_|-].*[\w|가-힣].+[_|-]MSPP[_|-][0-9]{4}[_|-][0-9]{2}[_|-][0-9]{2}[_|-][\w|가-힣|_].*");


                    Regex regex = new Regex(@"[\w|_|-]*[\w|가-힣].+[_|-]MSPP[_|-][0-9]{4}[_|-][0-9]{2}[_|-][0-9]{2}[_|-]*[\w|가-힣|_]*");

                    // 1. masterport-site-MSPP-####-##-##-description 형식인지 체크
                    if (regex.IsMatch(targetStr.Get(context)))
                    {
                        Regex splitNeNameRex = new Regex(@"[_|-]MSPP[_|-]");

                        // 2. MSPP 전/후로 나눔
                        string[] splitNeName = splitNeNameRex.Split(targetStr.Get(context));

                        string frontNeName = splitNeName[0];

                        string masterPort = frontNeName.Split('_')[0];
                        // 3. 마스터 포트가 존재하는경우 붙여서 저장
                        if (masterPort.Equals(""))
                        {
                            resultDic.Add("instlocation", frontNeName.Split('_')[1]);
                        }
                        else
                        {
                            resultDic.Add("instlocation", frontNeName);
                        }

                        // 4. 베이-쉘프-시스템 파싱하여 저장
                        string endNeName = splitNeName[1];

                        resultDic.Add("bay", endNeName.Split('_')[0]);
                        resultDic.Add("shelf", endNeName.Split('_')[1]);
                        resultDic.Add("system", endNeName.Split('_')[2]);

                        result.Set(context, resultDic);
                    }
                    else
                    {
                        Console.WriteLine("<<E>> TFSysnameParser - \"{0}\" 는(은) 파싱할 수 없는 포맷입니다.", targetStr.Get(context));
                    }



                }
            }

            /*
             * 텔레필드 EMS 등록정보중 cot rt ID를 tid로부터 파싱한다.
             * cot-id 10~98
             * rt-id 0~98
             */
            public class TFTidParser : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<String> targetStr { get; set; }

                [Category("Output")]
                public OutArgument<String> cotID { get; set; }

                [Category("Output")]
                public OutArgument<String> rtID { get; set; }

                [Category("Output")]
                public OutArgument<Boolean> succYN { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Boolean succYN = false;
                    String[] splitTid = this.targetStr.Get(context).Split('-');

                    if (splitTid.Count() >= 2)
                    {
                        int cot = 0;
                        Boolean cotNumbericCheck = int.TryParse(splitTid[0], out cot);

                        if (cotNumbericCheck == true && cot >= 10 && cot <= 98)
                        {
                            int rt = 0;
                            Boolean rtNumbericCheck = int.TryParse(splitTid[1], out rt);

                            if (rtNumbericCheck == true && rt >= 0 && rt <= 98)
                            {
                                succYN = true;
                                this.cotID.Set(context, splitTid[0]);
                                this.rtID.Set(context, splitTid[1]);
                            }
                        }
                    }

                    if (succYN == false)
                    {
                        this.cotID.Set(context, null);
                        this.rtID.Set(context, null);
                        Console.WriteLine("<<E>> TFParser - 파싱할 수 없는 형식입니다.");

                    }
                    this.succYN.Set(context, succYN);
                }

            }

            /*
             * EMS 자동등록시 해당건(row)에 대해 결과 여부를 이력 테이블에 업데이트 하는 기능 수행
             */
            public class HistUpdate : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<Dictionary<String, String>> dictionary { get; set; }

                [Category("Input"), RequiredArgument]
                public InArgument<DataSet> dataSet { get; set; }

                [Category("Input"), RequiredArgument]
                public InArgument<String> successYN { get; set; }
                [Category("Input")]
                public InArgument<String> cause { get; set; }
                [Category("Input")]
                public InArgument<String> etc { get; set; }

                [Category("Output")]
                public OutArgument<DataSet> resultSet { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> HistUpdate - Start!");
                    resultSet.Set(context, work(this.dictionary.Get(context), this.dataSet.Get(context), this.successYN.Get(context), this.cause.Get(context), this.etc.Get(context)));
                    Console.WriteLine("<<C>> HistUpdate - End!");
                }

                private DataSet work(Dictionary<String, String> dictionary, DataSet dataSet, String successYN, String cause, String etc)
                {
                    if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows == null || dataSet.Tables[0].Rows.Count == 0)
                    {
                        Console.WriteLine("<<E>> HistUpdate - EMS 자동 등록 이력 테이블셋이 null이거나 데이터테이블 관련 요소들을 찾을 수 없습니다.");
                        return null;
                    }
                    if (dictionary == null || dictionary.Count == 0 || successYN == null || !GlobalConstants.EMSRESULTCODELIST.Contains(successYN))
                    {
                        Console.WriteLine("<<E>> HistUpdate - 파라메터 오류");
                        return dataSet;
                    }

                    try
                    {

                        DateTime currentTime = DateTime.Now;
                        String debugMsg = JsonConvert.SerializeObject(dataSet.Tables[0]);
                        Console.WriteLine("Before : {0}", debugMsg);

                        foreach (DataRow row in dataSet.Tables[0].Rows)
                        {
                            if (row["vendor"].Equals("WR") && row["vendor"].Equals(dictionary["vendor"]))
                            {
                                // 제작사가 우리넷인 경우 neName과 pNeName과 같은 hist row를 찾아 이력을 업데이트 한다.
                                if (row["neName"].Equals(dictionary["neName"]) && row["pNeName"].Equals(dictionary["pNeName"]))
                                {
                                    row["regResult"] = successYN;
                                    if (cause != null)
                                    {
                                        row["regResultMsg"] = cause;
                                    }
                                    if (etc != null)
                                    {
                                        row["regResultCode"] = etc;
                                    }
                                    // 등록 시간 이력에 넘기기 위해 추가
                                    row["regdt"] = currentTime.ToString("yyyy-MM-dd HH:mm:ss");

                                    break;
                                }

                            }
                            else if (row["vendor"].Equals("CO") && row["vendor"].Equals(dictionary["vendor"]))
                            {
                                // 제작사가 코위버인 경우 로직 처리
                                if (row["neName"].Equals(dictionary["neName"]) && row["pNeName"].Equals(dictionary["pNeName"]))
                                {
                                    row["regResult"] = successYN;
                                    if (cause != null)
                                    {
                                        row["regResultMsg"] = cause;
                                    }
                                    if (etc != null)
                                    {
                                        row["regResultCode"] = etc;
                                    }
                                    // 등록 시간 이력에 넘기기 위해 추가
                                    row["regdt"] = currentTime.ToString("yyyy-MM-dd HH:mm:ss");

                                    break;
                                }
                            }
                            else if (row["vendor"].Equals("TF") && row["vendor"].Equals(dictionary["vendor"]))
                            {
                                // 제작사가 텔레필드인 경우 로직 처리
                                if (row["neName"].Equals(dictionary["neName"]))
                                {
                                    row["regResult"] = successYN;
                                    if (cause != null)
                                    {
                                        row["regResultMsg"] = cause;
                                    }
                                    if (etc != null)
                                    {
                                        row["regResultCode"] = etc;
                                    }
                                    // 등록 시간 이력에 넘기기 위해 추가
                                    row["regdt"] = currentTime.ToString("yyyy-MM-dd HH:mm:ss");

                                    break;
                                }

                            }
                        }

                        String debugMsg2 = JsonConvert.SerializeObject(dataSet.Tables[0]);
                        Console.WriteLine("After : {0}", debugMsg2);

                        return dataSet;
                    }
                    catch (Exception E)
                    {
                        Console.WriteLine("<<E>> HistUpdate - 예외 발생 \n    ExceptionType : {0}\n{1}", E.GetType(), E.Message.ToString());
                        return dataSet;
                    }

                }
            }

        } // [EMS] end of Biz

        namespace Converter
        {
            /*
             * DictionaryList를 DataSet형태로 반환
             */
            public class DictionaryListToDataSetActivity : CodeActivity
            {
                private List<Dictionary<String, String>> convertedDictionaryList { get; set; }
                private String configTableName { get; set; }
                private DataSet returnedDataSet { get; set; }

                [Category("Input")]
                [RequiredArgument]
                public InArgument<List<Dictionary<String, String>>> DictionaryList { get; set; }

                [Category("Option")]
                public InArgument<String> TableName { get; set; }

                [Category("Output")]
                public OutArgument<DataSet> DataSet { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> DictionaryListToDataSetActivity Start!");
                    setFiled(DictionaryList.Get(context), TableName.Get(context));
                    converterDictionaryList2DataSet();
                    DataSet.Set(context, returnedDataSet);
                    Console.WriteLine("<<C>> DictionaryListToDataSetActivity End");
                }

                public void setFiled(List<Dictionary<String, String>> dictionaryList, String tableName)
                {
                    this.convertedDictionaryList = dictionaryList;
                    this.configTableName = tableName;
                }

                public void converterDictionaryList2DataSet()
                {
                    try
                    {
                        activityParamValidCheck();
                        converter();
                    }
                    catch (Exception e)
                    {
                        logError(e);
                    }
                }

                private void activityParamValidCheck()
                {
                    if (convertedDictionaryList == null || convertedDictionaryList.Count == 0)
                    {
                        Console.WriteLine("파라메터 딕셔너리 리스트가 null이거나 사이즈가 0");
                        throw new Exception();
                    }
                    if (configTableName != null && configTableName.Equals(""))
                    {
                        Console.WriteLine("테이블명이 empty string 입니다.");
                        throw new Exception();
                    }

                }

                private void logError(Exception e)
                {
                    Console.WriteLine("<<E>> DictionaryListToDataSetActivity - 예외 / 타입 : {0}, 메세지 : \n{1}", e.GetType().ToString(), e.Message);
                }

                private void converter()
                {
                    returnedDataSet = new DataSet();

                    List<String> colums = new List<String>();

                    foreach (Dictionary<String, String> item in convertedDictionaryList)
                    {
                        foreach (String key in item.Keys)
                        {
                            if (!colums.Contains(key))
                            {
                                colums.Add(key);
                            }
                        }
                    }
                    DataTable resultDT = new DataTable();

                    if (configTableName != null)
                    {
                        resultDT.TableName = configTableName;
                    }

                    returnedDataSet.Tables.Add(resultDT);

                    // 2. 컬럼 설정
                    foreach (String col in colums)
                    {
                        resultDT.Columns.Add(col);
                    }

                    //2020.01.15 벤더 공통정보 반환하기 위해 추가 컬럼
                    List<String> extentionColList = new List<string>()
                    {
                        {"tid"},
                        {"site"},
                        {"sysName"},
                        {"equipType"},
                        {"ipAddr"}
                    };

                    foreach (String extentionCol in extentionColList)
                    {
                        if (!resultDT.Columns.Contains(extentionCol))
                        {
                            resultDT.Columns.Add(extentionCol);
                        }
                    }

                    resultDT.Columns.Add("regdt");
                    resultDT.Columns.Add("regResult");
                    resultDT.Columns.Add("regResultMsg");
                    resultDT.Columns.Add("regResultCode");

                    // 2-2. 컬럼 타입 설정
                    foreach (DataColumn col in resultDT.Columns)
                    {
                        col.DataType = System.Type.GetType("System.String");
                    }

                    // 3. row로 변환
                    foreach (Dictionary<String, String> item in convertedDictionaryList)
                    {
                        DataRow row = resultDT.NewRow();
                        foreach (DataColumn col in resultDT.Columns)
                        {
                            if (item.ContainsKey(col.ToString()))
                            {
                                if (item[col.ToString()] != null)
                                {
                                    row[col] = item[col.ToString()];
                                }
                            }
                            else
                            {
                                if (col.ToString().Equals("regResult"))
                                {
                                    row[col] = "-";
                                }
                                //row[col] = "-";
                            }


                        }
                        resultDT.Rows.Add(row);
                    }
                }

                /*private DataSet ConvertDictionaryListToDataTable(List<Dictionary<String, String>> dictionaryList, String tableName)
                {
                    if (dictionaryList == null || dictionaryList.Count == 0)
                    {
                        Console.WriteLine("<<E>> DictionaryListToDataSetActivity.ConvertDictionaryListToDataTable - 파라메터가 null 이거나 Size = 0");
                        return null;
                    }

                    DataSet resultDataset = new DataSet();
                    // 1. Key 추출
                    List<String> colums = new List<String>();

                    foreach (Dictionary<String, String> item in dictionaryList)
                    { 
                        foreach (String key in item.Keys)
                        {
                            if (!colums.Contains(key))
                            {
                                colums.Add(key);
                            }
                        }
                    }
                    DataTable resultDT = new DataTable();

                    if (tableName != null && !tableName.Equals(""))
                    {
                        resultDT.TableName = tableName;
                    }

                    resultDataset.Tables.Add(resultDT);

                    // 2. 컬럼 설정
                    foreach (String col in colums)
                    {
                        resultDT.Columns.Add(col);                       
                    }

                    //2020.01.15 벤더 공통정보 반환하기 위해 추가 컬럼
                    List<String> extentionColList = new List<string>()
                    {
                        {"tid"},
                        {"site"},
                        {"sysName"},
                        {"equipType"},
                        {"ipAddr"}
                    };

                    foreach(String extentionCol in extentionColList)
                    {
                        if (!resultDT.Columns.Contains(extentionCol))
                        {
                            resultDT.Columns.Add(extentionCol);
                        }
                    }

                    resultDT.Columns.Add("regdt");
                    resultDT.Columns.Add("regResult");
                    resultDT.Columns.Add("regResultMsg");
                    resultDT.Columns.Add("regResultCode");

                    // 2-2. 컬럼 타입 설정
                    foreach (DataColumn col in resultDT.Columns)
                    {
                        col.DataType = System.Type.GetType("System.String");
                    }

                    // 3. row로 변환
                    foreach (Dictionary<String, String> item in dictionaryList)
                    {
                        DataRow row = resultDT.NewRow();
                        foreach (DataColumn col in resultDT.Columns)
                        {
                            if (item.ContainsKey(col.ToString()))
                            {
                                if (item[col.ToString()] != null)
                                {
                                    row[col] = item[col.ToString()];
                                }
                            }
                            else
                            {
                                if (col.ToString().Equals("regResult"))
                                {
                                    row[col] = "-";
                                }
                                //row[col] = "-";
                            }
                            

                        }
                        resultDT.Rows.Add(row);
                    }

                    return resultDataset;
                }*/
            }

            /** 
             * =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
             * <p># @Project   : RPA_Controller
             * <p># @package   : -
             * <p># @Class     : DictionaryToJsonString
             * <p># @Extends   : -
             * <p># @Impl      : -
             * <p># @Date      : 2021.01.06
             * <p># @Author    : 강준영
             * <p># @Tag       : -
             * <p># @Desc      : Dictionary to JsonString
             * =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- */
            public class DictionaryToJsonString : CodeActivity
            {
                [Category("Input"), RequiredArgument, Description("Dictionary <String, String>")]
                public InArgument<Dictionary<String, String>> input { get; set; }

                [Category("Output"), Description("Json 포맷 형태의 String을 반환 합니다.")]
                public OutArgument<String> output { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    Dictionary<String, String> param = input.Get(context);
                    String tmpStr;

                    if (param != null)
                    {
                        Js.JObject tmp = Js.JObject.FromObject(param);
                        tmpStr = tmp.ToString();

                        output.Set(context, tmpStr);
                    }
                    else 
                    {
                        output.Set(context, "");
                    }
                    
                }

            }

            /** 
             * =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
             * <p># @Project   : RPA_Controller
             * <p># @package   : -
             * <p># @Class     : DictionaryListToJsonString
             * <p># @Extends   : -
             * <p># @Impl      : -
             * <p># @Date      : 2021.01.06
             * <p># @Author    : 강준영
             * <p># @Tag       : -
             * <p># @Desc      : DictionaryList<String,String> to JsonString
             * =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=- */
            public class DictionaryListToJsonString : CodeActivity
            {
                [Category("Input"), RequiredArgument, Description("List<Dictionary<String, String>>")]
                public InArgument<List<Dictionary<String, String>>> input { get; set; }

                [Category("Output"), Description("Json 포맷 형태의 String을 반환 합니다.")]
                public OutArgument<String> output { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    String tmpStr;
                    List<Dictionary<String, String>> param = input.Get(context);

                    if (param != null)
                    {
                        Js.JArray tmpList = Js.JArray.FromObject(param);
                        tmpStr = tmpList.ToString();

                        output.Set(context, tmpStr);
                    }
                    else
                    {
                        output.Set(context, "");
                    }
                }
            }

            /*
             * JsonString => List<Dictionary<String, String>> 으로 변환
             */
            public class JsonStringToDictionaryList : CodeActivity
            {
                private String convertedJson { get; set; }
                private List<Dictionary<String, String>> returnedDictionaryList { get; set; }
                private Boolean isOk { get; set; }

                [Category("Input"), RequiredArgument, Description("적절한 json포맷의 String값을 입력")]
                public InArgument<String> JsonString { get; set; }

                [Category("Output")]
                public OutArgument<List<Dictionary<String, String>>> resultList { get; set; }

                [Category("Output"), Description("정상 데이터 반환 여부를 확인합니다.")]
                public OutArgument<Boolean> outYN { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> JsonStringToDictionaryList - RPC호출 Json 파싱수행 Start!");
                    getActivityParam(context);
                    setActivityResult(context);
                    Console.WriteLine("<<C>> JsonStringToDictionaryList - End");
                }

                void setActivityResult(CodeActivityContext context)
                {
                    resultList.Set(context, convertJsonToDictionaryList());
                    outYN.Set(context, isOk);
                }

                void getActivityParam(CodeActivityContext context)
                {
                    convertedJson = JsonString.Get(context);
                }

                void logError(Exception e)
                {
                    Console.WriteLine("<<E>> JsonStringToDictionaryList - 예외 / 타입 : {0}, 메세지 : \n{1}", e.GetType().ToString(), e.Message);
                }

                void logInfo()
                {
                    if (isOk)
                    {
                        Console.WriteLine("<<C>> JsonStringToDictionaryList - Complete");
                        foreach (Dictionary<String, String> i in returnedDictionaryList)
                        {
                            String tmp = JsonConvert.SerializeObject(i);
                            Console.WriteLine(tmp);
                        }
                    }
                    else
                        Console.WriteLine("<<E>> JsonStringToDictionaryList - 결과가 NULL이거나, 리스트의 요소가 0 입니다.");
                }

                public List<Dictionary<String, String>> convertJsonToDictionaryList()
                {
                    try
                    {
                        isOk = false;
                        isValidActivityParam();
                        isRPCok();
                        returnedDictionaryList = new List<Dictionary<String, String>>();
                        converter();
                        isOk = true;
                        return returnedDictionaryList;
                    }
                    catch (Exception e)
                    {
                        logError(e);
                        return null;
                    }
                    finally
                    {
                        logInfo();
                    }
                }
                private Boolean isEmpty(Object o)
                {
                    if (o == null)
                        return true;
                    if (typeof(String).Equals(o) && ((String)o).Trim().Length == 0)
                        return true;
                    return false;
                }

                private void isValidActivityParam()
                {
                    if (isEmpty(convertedJson))
                    {
                        Console.WriteLine("<<E>> JsonStringToDictionaryList - RPC 요청 파라메터가 NULL이거나 잘못됨");
                        throw new ArgumentNullException();
                    }
                }

                private void isRPCok()
                {
                    Js.JObject RpcResponseJson = Js.JObject.Parse(convertedJson);
                    Js.JObject outputOfRpcResponseJson = (Js.JObject)RpcResponseJson["output"];
                    String rpcResultCode = outputOfRpcResponseJson["resultMsg"].ToString();
                    Js.JArray regInfoList = (Js.JArray)outputOfRpcResponseJson["emsRegList"];

                    if (rpcResultCode.Equals("SUCC"))
                    {
                        if (regInfoList.Count == 0)
                        {
                            Console.WriteLine("요청 결과 데이터 list size가 0 인경우. 등록 대상 없음.");
                            throw new Exception();
                        }
                        Console.WriteLine("RPC 요청결과 정상");
                    }
                    else
                    {
                        Console.WriteLine("RPC 요청 결과 실패");
                        throw new Exception();
                    }
                }

                private void converter()
                {
                    Js.JObject RpcResponseJson = Js.JObject.Parse(convertedJson);
                    Js.JObject outputOfRpcResponseJson = (Js.JObject)RpcResponseJson["output"];
                    Js.JArray regInfoList = (Js.JArray)outputOfRpcResponseJson["emsRegList"];
                    foreach (Js.JObject element in regInfoList)
                    {
                        String convertedToString = JsonConvert.SerializeObject(element);
                        Dictionary<String, String> convertedToDictionary = JsonConvert.DeserializeObject<Dictionary<String, String>>(convertedToString);
                        returnedDictionaryList.Add(convertedToDictionary);
                    }
                }
            }

            public class EmsRegHistToJson : CodeActivity
            {
                [Category("Input"), RequiredArgument]
                public InArgument<DataSet> dataSet { get; set; }

                [Category("Output"), RequiredArgument]
                public OutArgument<String> resultJsonStr { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> EmsRegHistToJson Start!");
                    this.resultJsonStr.Set(context, convertDataSet2Json(this.dataSet.Get(context)));
                    Console.WriteLine("<<C>> EmsRegHistToJson End!");
                }

                public String convertDataSet2Json(DataSet dataSet)
                {
                    try
                    {
                        return convertDataSet2Json4RpcRequest(dataSet);
                    }
                    catch (Exception e)
                    {
                        logError(e);
                        return null;
                    }
                }
                public void logError(Exception e)
                {
                    Console.WriteLine("<<E>> EmsRegHistToJson - 예외 / 타입 : {0}, 메세지 : \n{1}", e.GetType().ToString(), e.Message);
                }

                private String convertDataSet2Json4RpcRequest(DataSet dataSet)
                {
                    String returnStr = null;
                    if (dataSet == null || dataSet.Tables[0] == null || dataSet.Tables[0].Rows == null || dataSet.Tables[0].Rows.Count == 0)
                    {
                        Console.WriteLine("<<E>> EmsRegHistToJson - DataSet가 NULL 또는 데이터가 존재 하지 않습니다.");
                        return null;
                    }
                    DataTable datatable = dataSet.Tables[0];
                    //JsonConvert
                    String JarrayString = JsonConvert.SerializeObject(datatable);
                    Js.JObject TopObj = new Js.JObject();
                    Js.JObject Jobj = new Js.JObject();
                    Js.JArray HistJarray = (Js.JArray)JsonConvert.DeserializeObject(JarrayString);
                    Jobj.Add("sessionId", "e14e5614-2fdb-4945-aace-f6be6b7fdef5");
                    Jobj.Add("emsRegHistList", HistJarray);
                    TopObj.Add("input", Jobj);
                    returnStr = JsonConvert.SerializeObject(TopObj);
                    Console.WriteLine("EMS 이력 RPC INPUT MESSAGE 결과 :: \n{0}", returnStr);
                    return returnStr;
                }
            }
        } // [EMS] end of Converter

        namespace Collection
        {
            /*
             * 벤더 선택후 정렬
             */
            public class SelectVendorByDictionaryList : CodeActivity
            {

                private List<Dictionary<String, String>> AllVendorDictionaryList;
                private List<Dictionary<String, String>> selectdVendorDictionaryList;
                private List<Dictionary<String, String>> sortedDictionaryList;
                private String vendorName;
                [Category("Input"), RequiredArgument]
                public InArgument<List<Dictionary<String, String>>> dictionaryList { get; set; }

                [Category("Input"), RequiredArgument]
                public InArgument<String> vendor { get; set; }

                [Category("Output")]
                public OutArgument<List<Dictionary<String, String>>> resultList { get; set; }
                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> SelectVendorByDictionaryList - Start");
                    this.vendorName = vendor.Get(context);
                    this.AllVendorDictionaryList = dictionaryList.Get(context);
                    resultList.Set(context, getDictionaryListOfSelectedVendorAndSortEmsIP());
                    Console.WriteLine("<<C>> SelectVendorByDictionaryList - End");
                }



                public void logError(Exception e)
                {
                    Console.WriteLine("<<E>> SelectVendorByDictionaryList - 예외 / 타입 : {0}, 메세지 : \n{1}", e.GetType().ToString(), e.Message);
                }

                public List<Dictionary<String, String>> getDictionaryListOfSelectedVendorAndSortEmsIP()
                {
                    try
                    {
                        activityParamValidCheck();
                        initailizationSelectdVendorDictionaryList();
                        initailizationSortedDictionaryList();
                        logInfo();
                        return sortedDictionaryList;
                    }
                    catch (Exception e)
                    {
                        logError(e);
                        return null;
                    }

                }

                public Boolean isDictionaryList()
                {
                    if (AllVendorDictionaryList == null || AllVendorDictionaryList.Count == 0)
                        return false;
                    else
                        return true;
                }

                public Boolean isVendorName()
                {
                    if (vendorName == null || !GlobalConstants.EMSVENDORLIST.Contains(vendorName))
                        return false;
                    else
                        return true;
                }

                public void activityParamValidCheck()
                {
                    if (!isVendorName())
                    {
                        Console.WriteLine("<<E>> SelectVendorByDictionaryList - 벤더명이 null이거나 벤더 리스트에 포함되지 않음");
                        throw new Exception();
                    }
                    if (!isDictionaryList())
                    {
                        Console.WriteLine("<<E>> SelectVendorByDictionaryList - DictionaryList가 NULL이거나 사이즈가 0 입니다.");
                        throw new Exception();
                    }
                }

                private List<String> createEmsIpList()
                {

                    List<String> emsIpList = new List<String>();
                    foreach (Dictionary<String, String> i in selectdVendorDictionaryList)
                    {
                        String emsip = i["emsIp"];
                        if (!emsIpList.Contains(emsip))
                        {
                            emsIpList.Add(emsip);
                        }
                    }
                    return emsIpList;
                }

                private void initailizationSelectdVendorDictionaryList()
                {

                    selectdVendorDictionaryList = new List<Dictionary<string, string>>();

                    foreach (Dictionary<String, String> element in AllVendorDictionaryList)
                    {
                        if (element["vendor"].Equals(vendorName))
                        {
                            selectdVendorDictionaryList.Add(element);
                        }
                    }
                }

                private void initailizationSortedDictionaryList()
                {
                    sortedDictionaryList = new List<Dictionary<string, string>>();

                    List<String> emsipList = createEmsIpList();

                    foreach (String emsip in emsipList)
                    {
                        foreach (Dictionary<String, String> i in selectdVendorDictionaryList)
                        {
                            if (emsip.Equals(i["emsIp"]))
                            {
                                sortedDictionaryList.Add(i);
                            }
                        }
                    }

                }
                private void logInfo()
                {
                    String debugStr = JsonConvert.SerializeObject(sortedDictionaryList);
                    Console.WriteLine("<<Info>> SelectVendorByDictionaryList - result \n{0}", debugStr);
                }


            }

            /*
             * EMS 미등록 건 리스트 반환
             */
            public class SelectUnRegByHist : CodeActivity
            {
                private DataSet histDataSet { get; set; }
                private List<Dictionary<String, String>> convertedList { get; set; }

                [Category("Input"), RequiredArgument]
                public InArgument<DataSet> regHist { get; set; }

                [Category("Output")]
                public OutArgument<List<Dictionary<String, String>>> resultDictionaryList { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    Console.WriteLine("<<C>> SelectUnRegByHist - Start");
                    histDataSet = regHist.Get(context);

                    resultDictionaryList.Set(context, convertDataSetToDictionaryList());
                    Console.WriteLine("<<C>> SelectUnRegByHist - End");
                }
                // vailed check, 미등록건 발췌

                public void activityParamValidCheck()
                {
                    if (!isDataSet())
                    {
                        Console.WriteLine("<<E>> SelectUnRegByHist - 파라메터가 null이거나, 사이즈가 0 입니다.");
                        throw new Exception();
                    }
                }
                public Boolean isDataSet()
                {
                    if (histDataSet == null || histDataSet.Tables.Count == 0)
                        return false;
                    else
                        return true;
                }
                public void logError(Exception e)
                {
                    Console.WriteLine("<<E>> SelectUnRegByHist - 예외 / 타입 : {0}, 메세지 : \n{1}", e.GetType().ToString(), e.Message);
                }

                private void convertDataSetToUnRegDictionaryList()
                {
                    convertedList = new List<Dictionary<String, String>>();

                    DataTable datatable = histDataSet.Tables[0];
                    Console.WriteLine("불러온 테이블명 : {0}", datatable.TableName);

                    foreach (DataRow row in datatable.Rows)
                    {
                        if (!row["regResult"].ToString().Equals("-"))
                            continue;
                        Dictionary<String, String> element = new Dictionary<string, string>();
                        foreach (DataColumn col in datatable.Columns)
                        {
                            element.Add(col.ToString(), row[col].ToString());

                        }
                        convertedList.Add(element);
                    }

                    String debugMsg = JsonConvert.SerializeObject(convertedList);
                    Console.WriteLine("<<C>> SelectUnRegByHist - result : \n{0}", debugMsg);
                }

                public List<Dictionary<String, String>> convertDataSetToDictionaryList()
                {
                    try
                    {
                        activityParamValidCheck();
                        convertDataSetToUnRegDictionaryList();
                        return convertedList;
                    }
                    catch (Exception e)
                    {
                        logError(e);
                        return null;
                    }

                }
            }



            /*
             * 임시 수작업으로 등록제외하고 싶은 리스트를 처리하기 위해서 기능 생성
             */
            public class RegListFilter : CodeActivity {

                [Category("Input"), RequiredArgument]
                public InArgument<String[]> regList { get; set; }

                [Category("output")]
                public InOutArgument<List<Dictionary<String, String>>> resultDicList { get; set; }

                protected override void Execute(CodeActivityContext context)
                {
                    List<Dictionary<String, String>> tmpDicList = resultDicList.Get(context);
                    String[] tmplist = regList.Get(context);
                    if (tmplist != null && tmplist.Length > 0)
                    {
                        
                        List<Dictionary<String, String>> toRemoveList = new List<Dictionary<String, String>>();

                        foreach (String item in tmplist) {

                            foreach (Dictionary<String, String> row in tmpDicList) {
                                if (row["tid"].Equals(item)) {
                                    Console.WriteLine("등록 제외건 리스트에서 삭제 : {0}", row["tid"]);
                                    toRemoveList.Add(row);
                                    break;
                                }
                                    
                            }
                        }

                        tmpDicList.RemoveAll(toRemoveList.Contains);

                    }
                    else 
                    {
                        Console.WriteLine("<<E>> RegListFilter - 파라메터  null");
                    }

                    resultDicList.Set(context, tmpDicList);
                }

            }
        }

    }

}
