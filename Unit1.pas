unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.Mask, JvExMask, JvToolEdit,
  Winapi.ActiveX, Winapi.ShlObj, Winapi.PropSys,
  System.TypInfo, Winapi.PropKey, FilePropInfo;

type
  TForm1 = class(TForm)
    JvFilenameEdit1: TJvFilenameEdit;
    ListView1: TListView;
    procedure JvFilenameEdit1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure SetListViewGroup;
//    procedure ListViewAdd(const GroupName: string; pItem: PPropertyItem; Count: Integer; const Store: IPropertyStore);
    procedure ShowPropertyInformation(const FileName: string);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.SetListViewGroup;
var
  Group: TPropertyGroupItem;
  Item: TListGroup;
  Str: string;
begin
  for Group in PropKeyGroup do
  begin
    Item := ListView1.Groups.Add;
    if GetDisplayName(Group.Key^, Str) then
      Item.Header := Format('%s %s:%u', [Str, Group.Key.fmtid.ToString, Group.Key.pid])
    else
      Item.Header := Format('%s:%u', [Group.Key.fmtid.ToString, Group.Key.pid])
  end;
end;

//procedure TForm1.ListViewAdd(const GroupName: string; pItem: PPropertyItem; Count: Integer; const Store: IPropertyStore);
//var
//  Group: TListGroup;
//  Item: TListItem;
//  hR: HResult;
//  Value: TPropVariant;
//  StrValue: TPropVariant;
//begin
//  Group := ListView1.Groups.Add;
//  Group.Header := GroupName;
//  while Count > 0 do
//  begin
//    Dec(Count);
//    hR := Store.GetValue(pItem.Key^, Value);
//    if Succeeded(hR) then
//    begin
//      if (Value.vt > VT_NULL) then
//      begin
//        hR := PropVariantChangeType(StrValue, Value, 0, VT_LPWSTR);
//        if Succeeded(hR) then
//        begin
//          Item := ListView1.Items.Add;
//          Item.GroupID := Group.ID;
//          Item.Caption := pItem.Name;
//          Item.SubItems.Add(Value.ToString);
//          PropVariantClear(StrValue);
//        end;
//      end;
//      PropVariantClear(Value);
//    end;
//    Inc(pItem);
//  end;
//end;

procedure TForm1.ShowPropertyInformation(const FileName: string);
var
  hR: HResult;
  Str: string;
  Store: IPropertyStore;
  Count: Cardinal;
  StrBuf: array[0..255] of Char;
  I, iGroup, iItem: Integer;
  Key: TPropertyKey;
  Value: TPropVariant;
  Item: TListItem;
begin
  ListView1.Clear;

  hR := SHGetPropertyStoreFromParsingName(PChar(FileName), nil, GPS_DEFAULT, IPropertyStore, Store);

  if Succeeded(hR) then
  begin
    ListView1.Items.BeginUpdate;
    try
//      ListViewAdd('Media', @FilePropMedia, Length(FilePropMedia), Store);
//      ListViewAdd('Video', @FilePropVideo, Length(FilePropVideo), Store);
//      ListViewAdd('Audio', @FilePropAudio, Length(FilePropAudio), Store);

      hR := Store.GetCount(Count); // 取得屬性數量
      if Succeeded(hR) then
      begin
        for I := 0 to Count - 1 do
        begin
          hR := Store.GetAt(I, Key); // 取得屬性識別碼
          if Failed(hR) then
            Continue;

          hR := Store.GetValue(Key, Value); // 取得屬性的值
          if Failed(hR) then
            Continue;

          try
            // 這裡過濾掉空值，VT_EMPTY = 0 表示沒有資料，VT_NULL = 1 則是表示"無"的值
            // 這兩個都不會有有效的資訊，所以在只顯示屬性值的範例中不需要。
            if Value.vt < 2 then
              Continue;

            // 歸類群組
            iGroup := IndexOfGroup(Key, iItem);
            if (iGroup < 0) or (iItem < 0) then
              iGroup := Integer(PK_PropGroup_Advanced);

            // 能執行到此表示已取得屬性值，所以新增項目至清單
            Item := ListView1.Items.Add;
            Item.GroupID := ListView1.Groups.Items[iGroup].ID;
            Item.Caption := GetDisplayName(Key);

            Str := '';

            // 利用系統函數轉換屬性變數內容為字串
            hR := PropVariantToString(Value, StrBuf, Length(StrBuf));
            if Succeeded(hR) then
            begin
              Str := WideCharToString(StrBuf);
            end
            else
            begin
              // 回傳值若為 STRSAFE_E_INSUFFICIENT_BUFFER，則表示緩衝區不足
              // 可以重新分配更大的緩衝區，但這只是範例，不需要
              // STRSAFE_E_INSUFFICIENT_BUFFER 是屬於 Failed 值的範疇
              if hR = STRSAFE_E_INSUFFICIENT_BUFFER then
                Str := Format('<Type: %u, long>', [Value.vt])
              else
                Str := Format('<Type: %u>', [Value.vt])
            end;

            Item.SubItems.Add(Str);
            Item.SubItems.Add(Key.fmtid.ToString);
            Item.SubItems.Add(Key.pid.ToString);
          finally
            // 只要有取得 PropVariant 的值，就必須要有相應的 PropVariantClear。
            PropVariantClear(Value);
          end;

        end;
      end;
    finally
      ListView1.Items.EndUpdate;
    end;
    // 每當為介面 (IInterface) 指定一個變數就會引動 AddRef，參考計數就會加一，
    // 而變數設為 nil 就會 Release，參考計數就會減一，當參考計數為零時該介面會釋放。
    //
    // 在 C++ 會看到如 Store->Release(); 在 Delphi 中則會自動處理。
    //
    // 一般離開變數區塊後介面參考計數會自動減少，正常會自動釋放，
    // 如果不須提前釋放可不用指定 nil。
    Store := nil;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  SetListViewGroup;
end;

procedure TForm1.JvFilenameEdit1Change(Sender: TObject);
var
  FileName: string;
begin
  FileName := JvFilenameEdit1.FileName;
  if FileExists(FileName, False) then
    ShowPropertyInformation(FileName)
  else
    ListView1.Clear;
end;

initialization
  // 系統 COM 介面需要先進行初始化，參數 COINIT_APARTMENTTHREADED 為主執行續使用
  // 多執行續中若是用於主執行續以外的執行續則為 COINIT_MULTITHREADED。
  CoInitializeEx(nil, COINIT_APARTMENTTHREADED);

finalization
  // 當使用完 COM 介面後，在退出執行續前須釋放。
  CoUninitialize;

end.
