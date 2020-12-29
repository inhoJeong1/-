# PrintOrder
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

using Excel = Microsoft.Office.Interop.Excel;

namespace TDSM01.DSM
{
  pubilc partial class DSM005 : UserControlBase
  {
    // 부서 정보를 담고 있을 테이블 
    DataTable _dtDept = new DataTable();
    
    // 원가 요소를 담고 있을 테이블
    DataTable _dtBGTCODA = new DataTable();
    
    //Excel Export 시 부서명 출력을 위한 테이블 
    DataTable _dt_dept_list = new DataTable();
    
    //Excel Export 시 해당 파일의 로컬 주소를 담을 변수
    string _strOpenFilePath = string.Empty;
    
    string _strYear = string.Empty;
    
    //Excel Import 시 Excel to DataTable 을 위한 칼럼명과 엑셀 파일의 해당 값의 컬럼 위치 값
    Dictionart<int,string> _ExcelColNM = new Dictionary<int,String>()
    {
      {1,"BGTCODA"},
      {2,"BGTYEAR"},
      {5,"BGTMONTH"},
      {7,"REQ_TEMCD"},
      {9,"BGTSUM"},
      {12,"RMK"},
    };
    
    Public DSM005()
    {
      InitializeComponent();
      ComboBind();
      
      QueryResponse response = ZQxml.Default.ExecuteDataSet("PKG_TDSM0001.TDSM0001_SEARCH_DEPT");
      _dtDept = response.DataSet.Table[0];
      
      QueryResponse response2 = ZQxml.Default.ExecuteDataSet("PKG_TDSM0001.TDSM0001_SEARCH_BGTCODA");
      _dtBGTCODA = response2.DataSet.Table[0];
      
      _dt_dept_list.Columns.Add("VALUE");
    }
    
    Private void ComboBind()
    {
      // 콤보 박스에 년도에 해당하는 값만 바인딩
      Data dtYear = new DataTable();
      dtYear.Columns.Add("YEAR",typeof(string));
      for(int i = 0; i <= 100; i++)
      {
        DataRow dr = dtYear.NewRow();
        dr["YEAR"] = Convert.ToString(DateTime.Today.Year + (50 - i));
        dtYear.Rows.Add(dr);
      }
      ComboBind(lkuYear,"YEAR","YEAR");
      lkuYear.EditValue = Convert.ToString(DateTime.Today.Year);
    }
    
    Private void btnSearch_Clike(object sender,EventArgs e)
    {
      try
      {
        Cursor.Current = Cursor.WaitCursor;
        
        string[] arrDept = btnDept.Text.Split('-');
        string[] arrBGTCODA = btnBGTCODA.Text.Split('-');
        QueryParameterCollection parameters = new QueryParameterCollection();
        parameters.Add("USEDEPT",arrDept[0]);
        parameters.Add("BGTYEAR", lkuYear.EditValue);
        parameters.Add("BGTCODA",arrBGTCODA[0]);
        parameters.Add("LMCODE",rdoLMCODE.EditValue]);
        
        QueryResponse response = ZQxml.Default.ExecuteDataSet("PKG_TDSM0005.TDSM0005_PLAN_BUDGET_SELECT",parameters);
        DataTable dt = response.DataSet.Table[0];
        
        _strYear = lkuYear.EditValue.ToString();
        
        // 동적 그리드 레이아웃을 위한 과정
        Delete_Grid_Band();
        Create_Grid_Band(dt);
        
        grdMain.DataSource = dt;
      }
      catch(Exception ex)
      {
        Cursor.Current = Cursor.Default;
        throw ex;
      }
      finally
      {
        Cursor.Current = Cursor.Default;
      }
    }
    
    Private void Delete_Grid_Band()
    {
      // 그리드에 필수 요소를 제외한 초가 칼럼 요소들 삭제
      int ColNU = grvMain.Columns.Count;
      int BanNU = grvMain.Bands.Count;
      
      for(int i = 12; i < ColNU; i++)
        grvMain.Column.Remove(grvMain.Columns[12]);
        
      for(int i = 8; i < BanNU; i++)
        grvMain.Bands.Remove(grvMain.Bands[8]);
    }
    
    Private void Create_Grid_Band(DataTable dt)
    {
      // Oracle 칼럼 동적 리턴에 따라 칼럼 및 밴드 생성 및 
      string[] arrDept = btnDept.Text.Split('-');
      QueryParameterCollection parameters = new QueryParameterCollection();
      parameters.Add("USEDEPT",arrDept[0]);
      
      QueryResponse response = ZQxml.Default.ExecuteDataSet("PKG_TDSM0002.TDSM0002_TEAM_SEARCH",parameters);
      DataTable Deptdt = response.DataSet.Table[0];
      
      int loop = (dt.Columns.Count - 10) / 4;
      _dt_dept_list.Clear();
      
      for(int i = 0; i < loop; i++)
      {
        GridBand band = new GridBand();
        band.Children.AddBand("계획예산");
        band.Children.AddBand("청구");
        band.Children.AddBand("잔예산");
        band.Children.AddBand("비고");
        
        for(int j = 0; j < 4; j++)
        {
          BandedGridColumn col = new BandedGridColumn();
          col.FieldName = dt.Columns[(10 + j) + (i * 4)].ColumnName.ToString();
          col.Visible = true;
          //j == 3 비고에 해당하는 값 비고 제외한 칼럼은 숫자 관련 필드 이기 때문에 Summary 추가 및 DisplayFormat 변경
          if(j != 3)
          {
            col.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            col.SummaryItem.FieldName = col.FieldName.ToString();
            col.SummaryItem.DisplayFormat = "{0:N0}";
            
            col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
            col.DisplayFormat.FormatString ="{0:N0}";
            
          }
          
          // j == 0 || j == 3 는 입력을 위한 칼럼 이기때문에 배경 색 
          if(j == 0 || j == 3)
            col.AppearanceCell.BackColor = Color.FromArgb(255,224,192);
          else
            col.OptionsColumn.AllowEdit = false;
            
          string[] arrdept = dt.columns[11 + (i * 4)].ColumnName.ToString().Split('-');
          string strDept = arrdept[0] + "-" + (DeptDT.Select("POST_ORGN = '" + arrdept[0] + "'").CopyToDataTable()).Rows[0][1].ToString();
          
          DataRow row = _dt_dept_list.NewRow();
          row["VALUE"] = strDept;
          _dt_dept_list.Row.Add(row);
          
          band.Caption = strDept;
          grvMain.Columns.Add(col);
          band.Children[j].Columns.Add(grvMain.Columns[(12 + j) + (i * 4)]);
        }
        grvMain.Bands.Add(band);
        band.Width = 360;
      }
    }
    
    Private void btnDept_ButtonClick(object sender,DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
    {
      DSMCOD popup = new DSMCOD(_dtDept,"부서");
      popup.ShowDialog(this);
      
      btnDept.Text = popup;
    }
    
    Private void btnBGTCODA_ButtonClick(object sender,DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
    {
      DSMCOD popup = new DSMCOD(_dtBGTCODA,"원가요소");
      popup.ShowDialog(this);
      
      btnBGTCODA.Text = popup;
    }
    
    private void btnSave_Clike(object sender,EventArgs e)
    {
      DataTable dataTable = grdMain.DataSource as DataTable;
      DataRow[] row = dataTable.Select("CHK = 'Y'");
      if(row.Count() < 1)
      {
        MsgBox.Show(this,"선택된 행이 없습니다.","행 누락",MessageBoxButtons.OK,ImageKinds.Warnning);
        return;
      }
      if(MsgBox.Show(this,"저장하시겠습니까?","질의",MessageBoxButtons.YesNo,ImageKinds,Question) != DialogResult.Yes)return;
      
      //동적으로 생성된 칼럼에 따라 서로 다른 구조의 테이블 통일화를 위한 변환 과정
      DataTable dt = Convert_Save_Table(row.CopyToDataTable());
      try
      {
        Cursor.Current = Cursor.WaitCursor;
        SaveData(dt);
      }
      catch(Excepting ex)
      {
        Cursor.Current = Cursor.Default;
        throw ex;
      }
      finally
      {
        Cursor.Current = Cursor.Default;
      }
    }
    private void SaveData(DataTable dtSave,bool ExcelMode = false)
    {
      string[] paramNames = new string[] { "BGTCODA" , "BGTYEAR" , "BGTMONTH" , "REQ_TEMCD" , "BGTSUM" , "RMK" };
      strint[] arrDept = btnDept.Text.Spilt('-');
      
      QueryParameterCollection parameters = new QueryParameterCollection();
      parameters.AddArrayBindParameters(paramNames,dtSave);
      parameters.Add("USEDEPT",arrDept[0]);
      
      QueryResponse res = ZQxml.Default.ExecuteNonQuery("PKG_TDSM0005.TDSM0005_PLAN_BUDGET_IMPORT",parameters,QueryServiceTransactions.TxNone);
      var outputParameter = res.Parameters;
      DataTable MsgDt = Support.Get_MsgParamter(outputParameter);
      
      //ExcelMode시 해당 파일을 오픈 후 그 파일에 저장 행위에 대한 결과값 
      if(ExcelMode)
        OpenExcelFile(MsgDt);
      else
        GridMsgMode(MsgDt);
      
    }
    
    private void OpenExcelFile(DataTable MsgDt)
    {
      //WD.Utility At
      ExcelCell ec = new ExcelCell(this);
      if(ec.Open(_strOpenFilePath))
      {
        for(int i = 0; i < MsgDt.Rows.Count; i++)
        {
          if(MsgDt.Rows[i]["CODE"].ToString() == 0)
            ec.Workbook.Worksheets.ActiveWorksheet.Cells["P" + (5 + i)].Value = "OK";
          else
            ec.Workbook.Worksheets.ActiveWorksheet.Cells["P" + (5 + i)].Value = MsgDt.Rows[i]["MSG"].ToString();
        }
      }
      string fileFuttllName = ec.SaveFile(_strOpenFilePath,false);
      openFile(fileFuttllName);
    }
    
    private void GridMsgMode(DataTable MsgDt)
    {
      DataRow[] row = (grdMain.DataSource as DataTable).Select("CHK = 'Y'");
      for(int i = 0; i < row.Count() ; i ++)
      {
        row[i]["SEQ"] = MsgDt.Rows[i]["CODE"].ToString() == "0" ? "OK" : "ER";
        row[i]["MSG"] = MsgDt.Rows[i]["MSG"].ToString() == "0" ? "OK" : "ER";
        row[i]["CHK"] = MsgDt.Rows[i]["CODE"].ToString() == "0" ? "N" : "Y";
      }
    }
    
    private DataTable Convert_Save_Table(DataTable dt)
    {
      DataTable Save_Table = new DataTable();
      foreach(var strcolNM in _ExcelColNM.Values)
        Save_Table.Columns.Add(strcolNM);
        
      int loop = (dt.Columns.Count - 13) / 4;
      
      foreach(DataRow row in dt.Rows)
      {
        for(int i = 0; i < loop; i++)
        {
          DataRow rows = Save_Table.NewRow();
          
          string[] arrDeptNM = dt.Columns[11 + (i * 4)].ToString().Spilt('_');
          string strDeptNM = arrDeptNM[0];
          rows["BGTCODA"] = row["BGTCODA"].ToString(); 
          rows["BGTYEAR"] = _strYear;
          rows["BGTMONTH"] = row["BGTMONTH"].ToString(); 
          rows["REQ_TEMCD"] = strDeptNM;
          rows["BGTSUM"] = row[11 + (i * 4)].ToString();
          rows["RMK"] = row[14 + (i * 4)].ToString();
          
          if(rows["BGTSUM"].ToString() == string.Empty)
            continue;
          else
            Save_Table.Rows.Add(rows);
        }
      }
      return Save_Table;
    }
    
    
    private void btnPrint_Clike(object sender,EventArgs e)
    {
      try
      {
        Cursor.Current = Cursor.WaitCursor;
        Print_Excel();
      }
      catch(Excepting ex)
      {
        throw ex;
      }
      finally
      {
        Cursor.Current = Cursor.Default;
      }
    }
    
    
    private void Print_Excel()
    {
      string fileName = "DSM005.xlsx";
      
      Excel.Application excelApp = null;
      Excel._Workbook excelbook = null;
      Excel.Sheets excelSheets = null;
      Excel._Worksheet excelSheet = null;
      
      DataTable dtExcel = (grdMain.DataSource as DataTable).Copy();
      dtExcel.Columns.Remove("CHK");
      
      try
      {
        ExcelSupport es = new ExcelSupport();
        excelbook. es.OpenExcel(fileName.false);
        excelSheets = excelbook.Worksheets;
        excelSheet = (Microsoft.Office.Interop.Excel._Worksheet)excelSheets.get_Item(1);
        
        string start_cell = string.Empty;
        string fin_cell = string.Empty;
    
        int loop(dtExcel.Columns.Count - 10) / 4;
        
        excelSheet.Cells[2 , 3] = btnDept.EditValue;
        excelSheet.Cells[2 , 5] = lukYear.EditValue;
        
        fin_cell = GenerarteSequence(18 + ((loop -2) * 4));
        excelSheet.get_Range("L2 , "O5").Copy(excelSheet.get_Range("P2 , fin_cell + "5"));
        
        for(int i = 0; i < loop; i ++)
        {
          start_cell = GenerarteSequence( 15 + (i * 4));
          fin_cell = GenerarteSequence( 18 + (i * 4));
          
          excelSheet.Cells[3 , 12 + (i * 4)] = _dt_dept_list.Rows[i]["VALUE"].ToString();
          excelSheet.get_Range(fin_cell + "4").ColumnWidth = 18;
        }
        
        Excel.Range rng = excelSheet.Range["A5" + fin_cell + (dtExcel.Rows.Count + 4).ToString()];
        excelSheet.get_Range("A5", fin_cell + "5").Copy(excelSheet.get_Range("A6", fin_cell + "5"));
        
        object[,] only_Data = (object[,])rng.get_Value();
        int row = dtExcel.Rows.Count;
        int colmn = dtExcel.Columns.Count;
        object[,] data = new object[row,colmn];
        data = only_Data;
        
        excelSheet.Range["A1",fin_cell + "1"].Merge();
        
        for(int i = 0; i < dtExcel.Rows.Count ; i++)
        {
          for(int j = 0; j < dtExcel.Columns.Count - 2 ; j++)
          {
            data[1 + i, 1 + j] = dtExcel.Rows[i][j];
          }
        }
        
        rng.Value = data;
        excelBook.Application.Visible = true;
      }
      catch(Excepting ex)
      {
        throw ex;
      }
      finally
      {
        if(excelSheet != null)
        {
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
          excelSheet = null;
        }
        if(excelSheet != null)
        {
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
          excelSheet = null;
        }
        if(excelSheets != null)
        {
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets);
          excelSheets = null;
        }
        if(excelbook != null)
        {
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelbook);
          excelbook = null;
        }
        if(excelApp != null)
        {
          System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
          excelApp = null;
        }
      }
    }
    
    private string GenerarteSequence(int num)
    {
      string str ="";
      char achar;
      int mod;
      
      while(true)
      {
        mod = (num % 26) 65;
        num = (int)(num /26);
        achar = (char)mod;
        str = achar + str;
        if(num > 0) num--;
        else if(num == 0) break;
      }
      
      retrun str;
    }
    
    
    private void btnIMPORT_Clike(object sender,EventArgs e)
    {
      try
      {
        Cursor.Current = Cursor.WaitCursor;
        DataTalbe dt = ExcelSupport.ExcelToDataTable(_ExcelColNM , 4);
        
        dt.AsEnumerable().ToList().ForEach(s => s["BGTYEAR"] = _strYear);
        this._strOpenFilePath = ExcelSupport._strOpenFilePath;
        SaveData(dt,true);
      }
      catch(Excepting ex)
      {
        throw ex;
      }
      finally
      {
        Cursor.Current = Cursor.Default;
      }
    }
    
  }
}

