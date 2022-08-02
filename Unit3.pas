unit Unit3;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.ExtCtrls,ComObj;

type
  TForm3 = class(TForm)
    Panel1: TPanel;
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    StringGrid1: TStringGrid;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

  function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
  const
    xlCellTypeLastCell = $0000000B;
  var
    XLApp, Sheet: OLEVariant;
    RangeMatrix: Variant;
    x, y, k, r: Integer;
  begin
    Result := False;
    // Create Excel-OLE Object
    XLApp := CreateOleObject('Excel.Application');
    try
      // Hide Excel
      XLApp.Visible := False;

      // Open the Workbook
      XLApp.Workbooks.Open(AXLSFile);

      // Sheet := XLApp.Workbooks[1].WorkSheets[1];
      Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];

      // In order to know the dimension of the WorkSheet, i.e the number of rows
      // and the number of columns, we activate the last non-empty cell of it

      Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
      // Get the value of the last row
      x := XLApp.ActiveCell.Row;
      // Get the value of the last column
      y := XLApp.ActiveCell.Column;

      // Set Stringgrid's row &col dimensions.

      AGrid.RowCount := x;
      AGrid.ColCount := y;

      // Assign the Variant associated with the WorkSheet to the Delphi Variant

      RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
      //  Define the loop for filling in the TStringGrid
      k := 1;
      repeat
        for r := 1 to y do
          AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
        Inc(k, 1);
        AGrid.RowCount := k + 1;
      until k > x;
      // Unassign the Delphi Variant Matrix
      RangeMatrix := Unassigned;

    finally
      // Quit Excel
      if not VarIsEmpty(XLApp) then
      begin
        // XLApp.DisplayAlerts := False;
        XLApp.Quit;
        XLAPP := Unassigned;
        Sheet := Unassigned;
        Result := True;
      end;
    end;
  end;

  procedure TForm3.Button1Click(Sender: TObject);
  begin
    if Xls_To_StringGrid(StringGrid1, 'C:\Table1.xls') then
      ShowMessage('Table has been exported!');
  end;


end.
