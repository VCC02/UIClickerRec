{
    Copyright (C) 2025 VCC
    creation date: 02 Sep 2025
    initial release date: 02 Sep 2025

    author: VCC
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"),
    to deal in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies of the Software, and to permit persons to whom the
    Software is furnished to do so, subject to the following conditions:
    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
    OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
}


unit UIClickerRecMainForm;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Spin, VirtualTrees, ClickerUtils, ImgList;

type
  TExtraActionDetails = record
    ControlHandle: THandle;
    ClickPoint: TPoint;
    Timestamp: TDateTime;
    ScrShot: TBitmap;
  end;

  TExtraActionDetailsArr = array of TExtraActionDetails;

  TRecording = record
    Actions: TClkActionsRecArr;        //These two arrays should be kepts in sync. An aray of TExtraActionDetails, containg a TClkActionsRec field, could be made, but there are already useful functions, written for TClkActionsRecArr.
    Details: TExtraActionDetailsArr;
  end;

  TRecordingArr = array of TRecording;

  { TfrmUIClickerRecMain }

  TfrmUIClickerRecMain = class(TForm)
    btnCopySelectedActionsToClipboard: TButton;
    btnClearRecording: TButton;
    chkIncludeThisRecorder: TCheckBox;
    chkUnconditionalScreenshots: TCheckBox;
    chkRec: TCheckBox;
    chkMouseMoveScreenshots1: TCheckBox;
    grpMouseActions: TGroupBox;
    grpMouseButtonStates: TGroupBox;
    imgPreview: TImage;
    imglstActions16: TImageList;
    lblMs1: TLabel;
    lblMs2: TLabel;
    memLog: TMemo;
    pnlLeft: TPanel;
    pnlMouseMove: TPanel;
    pnlMiddle: TPanel;
    pnlMouseDrag: TPanel;
    pnlRight: TPanel;
    shpRec: TShape;
    spnedtUnconditionalPeriod: TSpinEdit;
    spnedtMouseMovePeriod: TSpinEdit;
    tmrMouseMoveDebounce: TTimer;
    tmrBlinkRec: TTimer;
    tmrRec: TTimer;
    vstRec: TVirtualStringTree;
    procedure btnClearRecordingClick(Sender: TObject);
    procedure btnCopySelectedActionsToClipboardClick(Sender: TObject);
    procedure chkRecChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure tmrBlinkRecTimer(Sender: TObject);
    procedure tmrMouseMoveDebounceTimer(Sender: TObject);
    procedure tmrRecTimer(Sender: TObject);
    procedure vstRecGetImageIndex(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Kind: TVTImageKind; Column: TColumnIndex; var Ghosted: boolean;
      var ImageIndex: integer);
    procedure vstRecGetImageIndexEx(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
      var Ghosted: boolean; var ImageIndex: integer;
      var ImageList: TCustomImageList);
    procedure vstRecGetText(Sender: TBaseVirtualTree; Node: PVirtualNode;
      Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure vstRecKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure vstRecMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure vstRecMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    FLeftButtonDown: Boolean;
    FRightButtonDown: Boolean;
    FMiddleButtonDown: Boolean;

    FLeftButtonUp: Boolean;
    FRightButtonUp: Boolean;
    FMiddleButtonUp: Boolean;

    FPrevLeftButtonDown: Boolean;
    FPrevRightButtonDown: Boolean;
    FPrevMiddleButtonDown: Boolean;

    FPrevLeftButtonUp: Boolean;
    FPrevRightButtonUp: Boolean;
    FPrevMiddleButtonUp: Boolean;

    FMouseMoving: Boolean;
    FMouseDragging: Boolean;

    FLastPos: TPoint;   //any event
    FLastPosLeft: TPoint;  //Left button
    FLastPosRight: TPoint;  //Right button
    FLastPosMiddle: TPoint;  //Middle button

    FAllRecs: TRecordingArr;
    FHitTimeStamp: QWord;

    procedure AddToLog(s: string);
    procedure DisplayMouseButtonStates;
    function GetSelectedActions(var AActionsArr: TClkActionsRecArr): Integer;
    procedure CopySelectedActionsToClipboard;
    procedure AddFindControlAndClickActions(var ATree: TCompRecArr; AClickActionName: string; AMouseButton: TMouseButton; ACurrentNow: TDateTime);

    procedure HandleOnLeftButtonDown;
    procedure HandleOnRightButtonDown;
    procedure HandleOnMiddleButtonDown;
    procedure HandleOnLeftButtonUp;
    procedure HandleOnRightButtonUp;
    procedure HandleOnMiddleButtonUp;
  public

  end;

var
  frmUIClickerRecMain: TfrmUIClickerRecMain;

implementation

{$R *.frm}


uses
  BitmapProcessing, ClickerActionProperties, ClickerTemplates,
  Clipbrd;

{ TfrmUIClickerRecMain }


procedure TfrmUIClickerRecMain.AddToLog(s: string);
begin
  memLog.Lines.Add(DateTimeToStr(Now) + '  ' + s);
end;


procedure TfrmUIClickerRecMain.chkRecChange(Sender: TObject);
begin
  tmrRec.Enabled := chkRec.Checked;
  tmrBlinkRec.Enabled := chkRec.Checked;

  if not tmrBlinkRec.Enabled then
  begin
    shpRec.Brush.Color := $000040;

    FLeftButtonDown := False;
    FRightButtonDown := False;
    FMiddleButtonDown := False;

    DisplayMouseButtonStates;
  end;

  shpRec.Hint := BoolToStr(tmrRec.Enabled, '', 'not ') + 'recording';

  chkUnconditionalScreenshots.Enabled := tmrRec.Enabled;
  chkMouseMoveScreenshots1.Enabled := tmrRec.Enabled;
  spnedtUnconditionalPeriod.Enabled := tmrRec.Enabled;
  spnedtMouseMovePeriod.Enabled := tmrRec.Enabled;
end;


procedure TfrmUIClickerRecMain.btnClearRecordingClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to Length(FAllRecs[0].Details) - 1 do
    FreeAndNil(FAllRecs[0].Details[i].ScrShot);

  SetLength(FAllRecs[0].Actions, 0);
  SetLength(FAllRecs[0].Details, 0);
  vstRec.RootNodeCount := 0;
  vstRec.Repaint;
end;


function TfrmUIClickerRecMain.GetSelectedActions(var AActionsArr: TClkActionsRecArr): Integer;
var
  Node: PVirtualNode;
begin
  Result := 0;
  Node := vstRec.GetFirstSelected;

  if Node = nil then
    Exit;

  repeat
    if vstRec.Selected[Node] then
    begin
      SetLength(AActionsArr, Result + 1);
      CopyActionContent(FAllRecs[0].Actions[Node^.Index], AActionsArr[Result]);

      Inc(Result);
    end;

    Node := Node^.NextSibling;
  until Node = nil;
end;


procedure TfrmUIClickerRecMain.CopySelectedActionsToClipboard;
var
  AStringList: TStringList;   //much faster than T(Mem)IniFile
  ActionsToCopy: TClkActionsRecArr;
begin
  AStringList := TStringList.Create;
  try
    AStringList.LineBreak := #13#10;
    GetSelectedActions(ActionsToCopy);
    try
      SaveTemplateWithCustomActionsToStringList_V2(AStringList, ActionsToCopy, '', '');
      Clipboard.AsText := AStringList.Text;
    finally
      SetLength(ActionsToCopy, 0);
    end;
  finally
    AStringList.Free;
  end;
end;


procedure TfrmUIClickerRecMain.btnCopySelectedActionsToClipboardClick(
  Sender: TObject);
begin
  CopySelectedActionsToClipboard;
end;


procedure TfrmUIClickerRecMain.FormCreate(Sender: TObject);
begin
  FLeftButtonDown := False;
  FRightButtonDown := False;
  FMiddleButtonDown := False;

  FPrevLeftButtonDown := False;
  FPrevRightButtonDown := False;
  FPrevMiddleButtonDown := False;

  FLeftButtonUp := False;
  FRightButtonUp := False;
  FMiddleButtonUp := False;

  FPrevLeftButtonUp := True;
  FPrevRightButtonUp := True;
  FPrevMiddleButtonUp := True;

  FMouseMoving := False;
  FMouseDragging := False;

  FLastPosLeft.X := -1;
  FLastPosLeft.Y := -1;
  FLastPosRight.X := -1;
  FLastPosRight.Y := -1;
  FLastPosMiddle.X := -1;
  FLastPosMiddle.Y := -1;

  SetLength(FAllRecs, 1);  //this should be 0, to start with an empty list of recordings
  SetLength(FAllRecs[0].Actions, 0);
  SetLength(FAllRecs[0].Details, 0);
end;


procedure TfrmUIClickerRecMain.tmrBlinkRecTimer(Sender: TObject);
begin
  tmrBlinkRec.Tag := tmrBlinkRec.Tag + 1;

  if tmrBlinkRec.Tag and 1 = 1 then
    shpRec.Brush.Color := $4040FF
  else
    shpRec.Brush.Color := $000040;
end;


procedure TfrmUIClickerRecMain.tmrMouseMoveDebounceTimer(Sender: TObject);
begin
  tmrMouseMoveDebounce.Enabled := False;
  FMouseMoving := False;
end;


procedure TfrmUIClickerRecMain.tmrRecTimer(Sender: TObject);
var
  tp: TPoint;
begin
  FLeftButtonDown := GetAsyncKeyState(VK_LBUTTON) < 0;
  FRightButtonDown := GetAsyncKeyState(VK_RBUTTON) < 0;
  FMiddleButtonDown := GetAsyncKeyState(VK_MBUTTON) < 0;

  FLeftButtonUp := not FLeftButtonDown;
  FRightButtonUp := not FRightButtonDown;
  FMiddleButtonUp := not FMiddleButtonDown;

  GetCursorPos(tp);

  if (FLastPos.X <> tp.X) or (FLastPos.Y <> tp.Y) then
    FMouseMoving := True
  else
    if FMouseMoving then
      tmrMouseMoveDebounce.Enabled := True;

  FMouseDragging := (FLeftButtonDown or FRightButtonDown or FMiddleButtonDown) and FMouseMoving;

  //Transitions detection
  if not FPrevLeftButtonDown and FLeftButtonDown then
    HandleOnLeftButtonDown;

  if not FPrevRightButtonDown and FRightButtonDown then
    HandleOnRightButtonDown;

  if not FPrevMiddleButtonDown and FMiddleButtonDown then
    HandleOnMiddleButtonDown;

  if not FPrevLeftButtonUp and FLeftButtonUp then
    HandleOnLeftButtonUp;

  if not FPrevRightButtonUp and FRightButtonUp then
    HandleOnRightButtonUp;

  if not FPrevMiddleButtonUp and FMiddleButtonUp then
    HandleOnMiddleButtonUp;

  FPrevLeftButtonDown := FLeftButtonDown;
  FPrevRightButtonDown := FRightButtonDown;
  FPrevMiddleButtonDown := FMiddleButtonDown;

  FPrevLeftButtonUp := FLeftButtonUp;
  FPrevRightButtonUp := FRightButtonUp;
  FPrevMiddleButtonUp := FMiddleButtonUp;

  FLastPos := tp;
  DisplayMouseButtonStates;
end;


procedure TfrmUIClickerRecMain.vstRecGetImageIndex(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
  var Ghosted: boolean; var ImageIndex: integer);
begin
  if Column <> 1 then
    Exit;

  ImageIndex := Ord(FAllRecs[0].Actions[Node^.Index].ActionOptions.Action);
end;


procedure TfrmUIClickerRecMain.vstRecGetImageIndexEx(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
  var Ghosted: boolean; var ImageIndex: integer; var ImageList: TCustomImageList);
begin
  if Column = 1 then
  begin
    ImageIndex := Ord(FAllRecs[0].Actions[Node^.Index].ActionOptions.Action);
    ImageList := imglstActions16;
  end;
end;


procedure TfrmUIClickerRecMain.vstRecGetText(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType;
  var CellText: string);
begin
  case Column of
    0: CellText := IntToStr(Node^.Index);
    1: CellText := FAllRecs[0].Actions[Node^.Index].ActionOptions.ActionName;
    2: CellText := CClkActionStr[FAllRecs[0].Actions[Node^.Index].ActionOptions.Action];
    3: CellText := FAllRecs[0].Actions[Node^.Index].FindControlOptions.MatchClassName;

    4:
    begin
      case FAllRecs[0].Actions[Node^.Index].ActionOptions.Action of
        acClick:
          CellText := FAllRecs[0].Actions[Node^.Index].ActionOptions.ActionName;

        acFindControl:
          CellText := FAllRecs[0].Actions[Node^.Index].FindControlOptions.MatchText;

        else
          CellText := 'Unimplemented';
      end;
    end;

    5: CellText := IntToStr(FAllRecs[0].Details[Node^.Index].ControlHandle);
    6: CellText := DateTimeToStr(FAllRecs[0].Details[Node^.Index].Timestamp);
    7: CellText := IntToStr(FAllRecs[0].Details[Node^.Index].ClickPoint.X) + ' : ' + IntToStr(FAllRecs[0].Details[Node^.Index].ClickPoint.Y);
  end;
end;


procedure TfrmUIClickerRecMain.vstRecKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ssCtrl in Shift then
  begin
    case Key of
      Ord('C'):
        CopySelectedActionsToClipboard;

    end;
  end;
end;


procedure TfrmUIClickerRecMain.vstRecMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FHitTimeStamp := GetTickCount64; //required on MouseUp
end;


procedure TfrmUIClickerRecMain.vstRecMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  Node: PVirtualNode;
begin
  if GetTickCount64 - FHitTimeStamp < 500 then  //This check is required, when double-clicking a file, in an OpenDialog, over vstRec.
  begin;                                        //The dialog closes at the second MouseDown event of the double-click action, while the second MouseUp triggers this (vstRecMouseUp) handler.
    Node := vstRec.GetFirstSelected;
    if Node = nil then
      Exit;

    try
      imgPreview.Picture.Bitmap.Assign(FAllRecs[0].Details[Node^.Index].ScrShot);
    except
    end;
  end;
end;


procedure TfrmUIClickerRecMain.DisplayMouseButtonStates;
begin
  if FLeftButtonDown then
    pnlLeft.Color := clYellow
  else
    pnlLeft.Color := clOlive;

  if FMiddleButtonDown then
    pnlMiddle.Color := clYellow
  else
    pnlMiddle.Color := clOlive;

  if FRightButtonDown then
    pnlRight.Color := clYellow
  else
    pnlRight.Color := clOlive;

  if FMouseMoving then
    pnlMouseMove.Color := clYellow
  else
    pnlMouseMove.Color := clOlive;

  if FMouseDragging then
    pnlMouseDrag.Color := clYellow
  else
    pnlMouseDrag.Color := clOlive;
end;


function GetComponentInfoAsStringAtPoint(APoint: TPoint): string;
var
  Comp: TCompRec;
begin
  Comp := GetWindowClassRec(APoint);
  Result := IntToStr(APoint.X) + ' : ' + IntToStr(APoint.Y) + ' (' + Comp.ClassName + ' / ' + Comp.Text + ' / ' + IntToStr(Comp.Handle) + ')';
  Result := FastReplace_0To1(Result);
end;


//function GetPreviousRecordedHandle(var ADetails: TExtraActionDetailsArr; ACurrentIndex: Integer): THandle;
//var
//  i: Integer;
//begin
//  Result := 0;
//  if (ACurrentIndex <= 0) or (ACurrentIndex > Length(ADetails) - 1) then
//    Exit;
//
//  Result := ADetails[ACurrentIndex - 1].ControlHandle;
//end;


function GetLastRecordedHandle(var ADetails: TExtraActionDetailsArr): THandle;
begin
  Result := 0;
  if Length(ADetails) = 0 then
    Exit;

  Result := ADetails[Length(ADetails) - 1].ControlHandle;
end;


procedure GetControlParentTree(AComp: TCompRec; var ATree: TCompRecArr);
var
  CompL, CompT, CompR, CompB: TCompRec;
  TargetProcID, ProcIDRes, ProcIDL, ProcIDT, ProcIDR, ProcIDB: DWord;
  TpL, TpU, TpR, TpB: TPoint;
  TreeLen: Integer;
  Found: Boolean;
begin
  if AComp.Handle = 0 then
    Exit;

  {$IFDEF Windows}
    ProcIDRes := GetWindowThreadProcessId(AComp.Handle, @TargetProcID);
    if ProcIDRes = 0 then
    begin
      //AddToLog
      Exit;
    end;
  {$ELSE}
    Exit;
  {$ENDIF}

  SetLength(ATree, 1);
  ATree[0] := AComp;

  if FindWindow(PChar(AComp.ClassName), PChar(AComp.Text)) > 0 then
    Exit;

  TreeLen := 0;
  repeat
    TpL.X := ATree[Length(ATree) - 1].ComponentRectangle.Left - 1;
    TpL.Y := ATree[Length(ATree) - 1].ComponentRectangle.Top;

    TpU.X := ATree[Length(ATree) - 1].ComponentRectangle.Left;
    TpU.Y := ATree[Length(ATree) - 1].ComponentRectangle.Top - 1;

    TpR.X := ATree[Length(ATree) - 1].ComponentRectangle.Right + 1;
    TpR.Y := ATree[Length(ATree) - 1].ComponentRectangle.Top;

    TpB.X := ATree[Length(ATree) - 1].ComponentRectangle.Left;
    TpB.Y := ATree[Length(ATree) - 1].ComponentRectangle.Bottom + 1;

    CompL := GetWindowClassRec(TpL);
    CompT := GetWindowClassRec(TpU);
    CompR := GetWindowClassRec(TpR);
    CompB := GetWindowClassRec(TpB);

    GetWindowThreadProcessId(CompL.Handle, @ProcIDL);
    GetWindowThreadProcessId(CompT.Handle, @ProcIDT);
    GetWindowThreadProcessId(CompR.Handle, @ProcIDR);
    GetWindowThreadProcessId(CompB.Handle, @ProcIDB);

    Found := False;
    if (ProcIDL = TargetProcID) and (CompL.Handle > 0) then
    begin
      SetLength(ATree, Length(ATree) + 1);
      ATree[Length(ATree) - 1] := CompL;
      Inc(TreeLen);
      Found := True;

      if CompL.Text > '' then  //There are windows with no caption (title), but they should be handled manually.
        if FindWindow(PChar(CompL.ClassName), PChar(CompL.Text)) > 0 then
          Break;
    end;

    if (ProcIDT = TargetProcID) and (CompT.Handle > 0) and not Found then
    begin
      SetLength(ATree, Length(ATree) + 1);
      ATree[Length(ATree) - 1] := CompT;
      Inc(TreeLen);
      Found := True;

      if CompT.Text > '' then  //There are windows with no caption (title), but they should be handled manually.
        if FindWindow(PChar(CompT.ClassName), PChar(CompT.Text)) > 0 then
          Break;
    end;

    if (ProcIDR = TargetProcID) and (CompR.Handle > 0) and not Found then
    begin
      SetLength(ATree, Length(ATree) + 1);
      ATree[Length(ATree) - 1] := CompR;
      Inc(TreeLen);
      Found := True;

      if CompR.Text > '' then  //There are windows with no caption (title), but they should be handled manually.
        if FindWindow(PChar(CompR.ClassName), PChar(CompR.Text)) > 0 then
          Break;
    end;

    if (ProcIDB = TargetProcID) and (CompB.Handle > 0) and not Found then
    begin
      SetLength(ATree, Length(ATree) + 1);
      ATree[Length(ATree) - 1] := CompB;
      Inc(TreeLen);
      Found := True;

      if CompB.Text > '' then  //There are windows with no caption (title), but they should be handled manually.
        if FindWindow(PChar(CompB.ClassName), PChar(CompB.Text)) > 0 then
          Break;
    end;

    if not Found then
      Break;
  until TreeLen > 100;
end;


procedure TfrmUIClickerRecMain.AddFindControlAndClickActions(var ATree: TCompRecArr; AClickActionName: string; AMouseButton: TMouseButton; ACurrentNow: TDateTime);
var
  i: Integer;
begin
  for i := Length(ATree) - 1 downto 0 do
    //if CompUp.Handle <> GetLastRecordedHandle(FAllRecs[0].Details) then  //makes no sense to check this, when starting from the top of the tree
    begin
      SetLength(FAllRecs[0].Actions, Length(FAllRecs[0].Actions) + 1);
      SetLength(FAllRecs[0].Details, Length(FAllRecs[0].Details) + 1);
      FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].Timestamp := ACurrentNow;
      FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ClickPoint := FLastPosLeft; //the MouseDown event
      FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ControlHandle := ATree[i].Handle;
      FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot := TBitmap.Create;
      ScreenShot(ATree[i].Handle, FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot, 0, 0, ATree[i].ComponentRectangle.Width, ATree[i].ComponentRectangle.Height);
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.Action := acFindControl;
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.ActionName := 'FindControl';
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.ActionEnabled := True;
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.ActionTimeout := 3000;

      GetDefaultPropertyValues_FindControl(FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].FindControlOptions);
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].FindControlOptions.MatchText := ATree[i].Text;
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].FindControlOptions.MatchClassName := ATree[i].ClassName;
      FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].FindControlOptions.UseWholeScreen := i = Length(ATree) - 1;
    end;

  //Click
  SetLength(FAllRecs[0].Actions, Length(FAllRecs[0].Actions) + 1);
  SetLength(FAllRecs[0].Details, Length(FAllRecs[0].Details) + 1);
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].Timestamp := ACurrentNow;
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ClickPoint := FLastPosLeft; //the MouseDown event
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ControlHandle := ATree[0].Handle;
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot := TBitmap.Create;
  ScreenShot(ATree[0].Handle, FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot, 0, 0, ATree[0].ComponentRectangle.Width, ATree[0].ComponentRectangle.Height);
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot.Canvas.Pen.Color := clRed;
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot.Canvas.Line(ATree[0].MouseXOffset, 0, ATree[0].MouseXOffset, ATree[0].ComponentRectangle.Height); //V
  FAllRecs[0].Details[Length(FAllRecs[0].Details) - 1].ScrShot.Canvas.Line(0, ATree[0].MouseYOffset, ATree[0].ComponentRectangle.Width, ATree[0].MouseYOffset); //H
  FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.Action := acClick;
  FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.ActionName := AClickActionName;
  FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ActionOptions.ActionEnabled := True;

  GetDefaultPropertyValues_Click(FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ClickOptions);
  FAllRecs[0].Actions[Length(FAllRecs[0].Actions) - 1].ClickOptions.MouseButton := AMouseButton;

  vstRec.RootNodeCount := Length(FAllRecs[0].Actions);
  vstRec.ClearSelection;
  vstRec.Selected[vstRec.GetLast] := True;
  vstRec.Repaint;
end;


procedure TfrmUIClickerRecMain.HandleOnLeftButtonDown;
var
  CompUp: TCompRec;
  CurrentNow: TDateTime;
  Tree: TCompRecArr;
begin
  FLastPosLeft := FLastPos;
  AddToLog('Left down on ' + GetComponentInfoAsStringAtPoint(FLastPos));

  //Special handling of MouseDown, as MouseClick, on menus, which react on MouseDown, so there would be no MouseUp to define a click:
  CompUp := GetWindowClassRec(FLastPosLeft);

  if not chkIncludeThisRecorder.Checked and
    ((CompUp.Handle = Handle) or
     (CompUp.Handle = chkIncludeThisRecorder.Handle) or
     (CompUp.Handle = btnCopySelectedActionsToClipboard.Handle) or
     (CompUp.Handle = btnClearRecording.Handle) or
     (CompUp.Handle = chkRec.Handle) or
     (CompUp.Handle = vstRec.Handle)) then
    Exit;

  CurrentNow := Now;
  GetControlParentTree(CompUp, Tree);
  AddFindControlAndClickActions(Tree, 'Click', mbLeft, CurrentNow);
end;


procedure TfrmUIClickerRecMain.HandleOnRightButtonDown;
begin
  FLastPosRight := FLastPos;
  AddToLog('Right down on ' + GetComponentInfoAsStringAtPoint(FLastPos));
end;


procedure TfrmUIClickerRecMain.HandleOnMiddleButtonDown;
begin
  FLastPosMiddle := FLastPos;
  AddToLog('Middle down on ' + GetComponentInfoAsStringAtPoint(FLastPos));
end;


function IsMenu(AComp: TCompRec): Boolean;
begin
  Result := True;
  if not ((AComp.ClassName = '#32768') and (AComp.Text = '')) then //system menu  - unfortunately, there are other windows with this handle
    Exit; //other menus (classes and other info) should be added, in order to identify components which are destroyed on MouseDown

  Result := False;
end;


function IsSelfClosingComponent(AComp: TCompRec): Boolean;
begin
  Result := True;

  if IsMenu(AComp) then
    Exit;

  if not ((AComp.ClassName = 'Button') and ((AComp.Text = '&Yes') or (AComp.Text = '&No') or (AComp.Text = 'OK'))) then //MessageBox buttons
    Exit;              //ParentClass = '#32770

  if not ((AComp.ClassName = 'Button') and ((AComp.Text = '&Open') or (AComp.Text = '&Save') or (AComp.Text = 'Select Folder') or (AComp.Text = 'Cancel'))) then //Open/Save dialog buttons
    Exit;              //ParentClass = '#32770

  Result := False;
end;


procedure TfrmUIClickerRecMain.HandleOnLeftButtonUp;
var
  CompDown: TCompRec;
  CompUp: TCompRec;
  CurrentNow: TDateTime;
  Tree: TCompRecArr;
begin
  AddToLog('Left up on ' + GetComponentInfoAsStringAtPoint(FLastPos));

  CompDown := GetWindowClassRec(FLastPosLeft);
  CompUp := GetWindowClassRec(FLastPos);

  if IsSelfClosingComponent(CompDown) then
    Exit;

  if (CompUp.Handle = CompDown.Handle) and (Abs(FLastPos.X - FLastPosLeft.X) < 6) and (Abs(FLastPos.Y - FLastPosLeft.Y) < 6) then
  begin
    if not chkIncludeThisRecorder.Checked and
      ((CompUp.Handle = Handle) or
       (CompUp.Handle = chkIncludeThisRecorder.Handle) or
       (CompUp.Handle = btnCopySelectedActionsToClipboard.Handle) or
       (CompUp.Handle = btnClearRecording.Handle) or
       (CompUp.Handle = chkRec.Handle) or
       (CompUp.Handle = vstRec.Handle)) then
      Exit;

    CurrentNow := Now;
    GetControlParentTree(CompUp, Tree);
    AddFindControlAndClickActions(Tree, 'Click', mbLeft, CurrentNow);
  end;
end;


procedure TfrmUIClickerRecMain.HandleOnRightButtonUp;
var
  CompDown: TCompRec;
  CompUp: TCompRec;
  CurrentNow: TDateTime;
  Tree: TCompRecArr;
begin
  AddToLog('Right up on ' + GetComponentInfoAsStringAtPoint(FLastPos));

  CompDown := GetWindowClassRec(FLastPosRight);
  CompUp := GetWindowClassRec(FLastPos);

  //if IsMenu then  Exit;

  if (CompUp.ClassName = '#32768') and (CompUp.ComponentRectangle.Left = FLastPos.X) and (CompUp.ComponentRectangle.Top = FLastPos.Y) then //system menu
  begin                 //Handled only the case where the pop-up menu opens to the right-bottom of the mouse cursor.
    Dec(FLastPos.X);
    Dec(FLastPos.Y);
    CompUp := GetWindowClassRec(FLastPos);  //attempt to capture the component outside the menu

    Dec(FLastPosRight.X);
    Dec(FLastPosRight.Y);
    CompDown := GetWindowClassRec(FLastPosRight);
  end;

  if (CompUp.Handle = CompDown.Handle) and (Abs(FLastPos.X - FLastPosRight.X) < 6) and (Abs(FLastPos.Y - FLastPosRight.Y) < 6) then
  begin
    if not chkIncludeThisRecorder.Checked and
      ((CompUp.Handle = Handle) or
       (CompUp.Handle = chkIncludeThisRecorder.Handle) or
       (CompUp.Handle = btnCopySelectedActionsToClipboard.Handle) or
       (CompUp.Handle = btnClearRecording.Handle) or
       (CompUp.Handle = chkRec.Handle) or
       (CompUp.Handle = vstRec.Handle)) then
      Exit;

    CurrentNow := Now;
    GetControlParentTree(CompUp, Tree);
    AddFindControlAndClickActions(Tree, 'Right-Click', mbRight, CurrentNow);
  end;
end;


procedure TfrmUIClickerRecMain.HandleOnMiddleButtonUp;
var
  CompDown: TCompRec;
  CompUp: TCompRec;
  CurrentNow: TDateTime;
  Tree: TCompRecArr;
begin
  AddToLog('Middle up on ' + GetComponentInfoAsStringAtPoint(FLastPos));

  CompDown := GetWindowClassRec(FLastPosMiddle);
  CompUp := GetWindowClassRec(FLastPos);

  if (CompUp.Handle = CompDown.Handle) and (Abs(FLastPos.X - FLastPosMiddle.X) < 6) and (Abs(FLastPos.Y - FLastPosMiddle.Y) < 6) then
  begin
    if not chkIncludeThisRecorder.Checked and
      ((CompUp.Handle = Handle) or
       (CompUp.Handle = chkIncludeThisRecorder.Handle) or
       (CompUp.Handle = btnCopySelectedActionsToClipboard.Handle) or
       (CompUp.Handle = btnClearRecording.Handle) or
       (CompUp.Handle = chkRec.Handle) or
       (CompUp.Handle = vstRec.Handle)) then
      Exit;

    CurrentNow := Now;
    GetControlParentTree(CompUp, Tree);
    AddFindControlAndClickActions(Tree, 'Middle-Click', mbMiddle, CurrentNow);
  end;
end;


end.

