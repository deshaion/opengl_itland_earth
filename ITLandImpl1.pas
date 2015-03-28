unit ITLandImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, ITLandProj1_TLB, StdVcl, OpenGL;

type
  TITLand = class(TActiveForm, IITLand)
    procedure ActiveFormPaint(Sender: TObject);
    procedure ActiveFormCreate(Sender: TObject);
    procedure ActiveFormDestroy(Sender: TObject);
  private
    { Private declarations }
    FEvents: IITLandEvents;
    DC: HDC;//**********
    hrc: HGLRC;//**********
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
    procedure Init;//*
    procedure SetDCPixelFormat;//*
    procedure PrepareImage(bmap: string);//*
  protected
    { Protected declarations }
    procedure DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage); override;
    procedure EventSinkChanged(const EventSink: IUnknown); override;
    function Get_Active: WordBool; safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_AutoScroll: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    function Get_Caption: WideString; safecall;
    function Get_Color: OLE_COLOR; safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    function Get_Enabled: WordBool; safecall;
    function Get_Font: IFontDisp; safecall;
    function Get_HelpFile: WideString; safecall;
    function Get_KeyPreview: WordBool; safecall;
    function Get_PixelsPerInch: Integer; safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    function Get_Scaled: WordBool; safecall;
    function Get_ScreenSnap: WordBool; safecall;
    function Get_SnapBuffer: Integer; safecall;
    function Get_Visible: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    procedure Set_Visible(Value: WordBool); safecall;
  public
    { Public declarations }
    procedure Initialize; override;
  end;

const
  Earth = 1;
  Itland=2;

var
  Angle : GLfloat = 0;
  time : LongInt;
  implementation

uses ComObj, ComServ;

{$R *.DFM}

{ TITLand }

procedure TITLand.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_ITLandPage); }
end;

procedure TITLand.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as IITLandEvents;
  inherited EventSinkChanged(EventSink);
end;

procedure TITLand.Initialize;
begin
  inherited Initialize;
  OnActivate := ActivateEvent;
  OnClick := ClickEvent;
  OnCreate := CreateEvent;
  OnDblClick := DblClickEvent;
  OnDeactivate := DeactivateEvent;
  OnDestroy := DestroyEvent;
  OnKeyPress := KeyPressEvent;
  OnPaint := PaintEvent;
end;

function TITLand.Get_Active: WordBool;
begin
  Result := Active;
end;

function TITLand.Get_AlignDisabled: WordBool;
begin
  Result := AlignDisabled;
end;

function TITLand.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TITLand.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TITLand.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TITLand.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TITLand.Get_Color: OLE_COLOR;
begin
  Result := OLE_COLOR(Color);
end;

function TITLand.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TITLand.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TITLand.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TITLand.Get_Font: IFontDisp;
begin
  GetOleFont(Font, Result);
end;

function TITLand.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TITLand.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TITLand.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TITLand.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TITLand.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TITLand.Get_ScreenSnap: WordBool;
begin
  Result := ScreenSnap;
end;

function TITLand.Get_SnapBuffer: Integer;
begin
  Result := SnapBuffer;
end;

function TITLand.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TITLand.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TITLand._Set_Font(var Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TITLand.ActivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnActivate;
end;

procedure TITLand.ClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnClick;
end;

procedure TITLand.CreateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnCreate;
end;

procedure TITLand.DblClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDblClick;
end;

procedure TITLand.DeactivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDeactivate;
end;

procedure TITLand.DestroyEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDestroy;
end;

procedure TITLand.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  TempKey := Smallint(Key);
  if FEvents <> nil then FEvents.OnKeyPress(TempKey);
  Key := Char(TempKey);
end;

procedure TITLand.PaintEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnPaint;
end;

procedure TITLand.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TITLand.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TITLand.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TITLand.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TITLand.Set_Color(Value: OLE_COLOR);
begin
  Color := TColor(Value);
end;

procedure TITLand.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TITLand.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TITLand.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TITLand.Set_Font(const Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TITLand.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TITLand.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TITLand.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TITLand.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TITLand.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TITLand.Set_ScreenSnap(Value: WordBool);
begin
  ScreenSnap := Value;
end;

procedure TITLand.Set_SnapBuffer(Value: Integer);
begin
  SnapBuffer := Value;
end;

procedure TITLand.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

//======================================================================
//Подготовка текстуры}
procedure TITLand.PrepareImage(bmap: string);
type
  PPixelArray = ^TPixelArray;
  TPixelArray = array [0..0] of Byte;
var
  Bitmap : TBitmap;
  Data, DataA : PPixelArray;
  BMInfo : TBitmapInfo;
  I, ImageSize : Integer;
  Temp : Byte;
  MemDC : HDC;
begin
  Bitmap := TBitmap.Create;
  Bitmap.LoadFromFile (bmap);
  with BMinfo.bmiHeader do begin
    FillChar (BMInfo, SizeOf(BMInfo), 0);
    biSize := sizeof (TBitmapInfoHeader);
    biBitCount := 24;
    biWidth := Bitmap.Width;
    biHeight := Bitmap.Height;
    ImageSize := biWidth * biHeight;
    biPlanes := 1;
    biCompression := BI_RGB;
    MemDC := CreateCompatibleDC (0);
    GetMem (Data, ImageSize * 3);
    GetMem (DataA, ImageSize * 4);
    try
      GetDIBits (MemDC, Bitmap.Handle, 0, biHeight, Data, BMInfo, DIB_RGB_COLORS);
      For I := 0 to ImageSize - 1 do begin
          Temp := Data [I * 3];
          Data [I * 3] := Data [I * 3 + 2];
          Data [I * 3 + 2] := Temp;
      end;

      For I := 0 to ImageSize - 1 do begin
          DataA [I * 4] := Data [I * 3];
          DataA [I * 4 + 1] := Data [I * 3 + 1];
          DataA [I * 4 + 2] := Data [I * 3 + 2];
          If (Data [I * 3 + 2] > 50) and
             (Data [I * 3 + 1] < 200) and
             (Data [I * 3] < 200) and (pos('ear',bmap)>0)
             then DataA [I * 4 + 3] := 27
             else DataA [I * 4 + 3] := 255;
          If (Data [I * 3 + 2] > 200) and
             (Data [I * 3 + 1] > 200) and
             (Data [I * 3] > 200) and (pos('name',bmap)>0)
             then DataA [I * 4 + 3] := 0 ;
           end;

      glTexImage2d(GL_TEXTURE_2D, 0, 3, biWidth,
                   biHeight, 0, GL_RGBA, GL_UNSIGNED_BYTE, DataA);
     finally
      FreeMem (Data);
      FreeMem (DataA);
      DeleteDC (MemDC);
      Bitmap.Free;
   end;
  end;
end;

{=======================================================================
Инициализация}
procedure TITLand.Init;
const
 LightPos : Array [0..3] of GLFloat = (10.0, 10.0, 0.0, 1.0);
var
 Quadric : GLUquadricObj;
begin
 Quadric := gluNewQuadric;
 gluQuadricTexture (Quadric, TRUE);

 glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST);
 glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST);

 glEnable(GL_TEXTURE_2D);

 glBlendFunc (GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA);

 glNewList (Earth, GL_COMPILE);
   prepareImage ('c:\earth.bmp');
    glEnable (GL_BLEND);
    glEnable(GL_CULL_FACE);
    glCullFace(GL_FRONT);
    gluSphere (Quadric, 1.0, 24, 24);
    glCullFace(GL_BACK);
    gluSphere (Quadric, 1.0, 24, 24);
    glDisable(GL_CULL_FACE);
    glDisable (GL_BLEND);
 glEndList;

 glNewList (Itland, GL_COMPILE);
   prepareImage ('c:\name3.bmp');
    glEnable (GL_BLEND);
    glEnable(GL_CULL_FACE);
    glCullFace(GL_FRONT);
    gluSphere (Quadric, 1.4, 24, 24);
    glCullFace(GL_BACK);
    gluSphere (Quadric, 1.4, 24, 24);
    glDisable(GL_CULL_FACE);
    glDisable (GL_BLEND);
 glEndList;

 gluDeleteQuadric (Quadric);

 glEnable(GL_DEPTH_TEST);
end;
procedure TITLand.ActiveFormPaint(Sender: TObject);
var
  ps : TPaintStruct;
begin
  BeginPaint(Handle, ps);

  glClear( GL_COLOR_BUFFER_BIT or GL_DEPTH_BUFFER_BIT );

  glPushMatrix;
     glRotatef (-10, 0.0, 1.0, 0.0);
     glRotatef (Angle, 0.0, 0.0, 1.0);
     glCallList(Earth);
  glPopMatrix;

  glPushMatrix;
     glRotatef (-30, 0.0, 1.0, 0.0);
     glRotatef (Angle, 0.0, 0.0, 1.0);
     glCallList(Itland);
  glPopMatrix;

  SwapBuffers(DC);
  EndPaint(Handle, ps);

  Angle := Angle + 0.25 * (GetTickCount - time) * 360 / 1000;
  If Angle >= 360.0 then Angle := 0.0;
  time := GetTickCount;

  glViewport(0, 0, ClientWidth, ClientHeight );
 glMatrixMode( GL_PROJECTION );
 glLoadIdentity;
 glFrustum( -1.0, 1.0, -1.0, 1.0, 5.0, 50.0 );
 glMatrixMode( GL_MODELVIEW );
 glLoadIdentity;
 glTranslatef( 0.0, 0.0, -12.0 );
 glRotatef(-90.0, 1.0, 0.0, 0.0);
  InvalidateRect(Handle, nil, False);
end;
{=======================================================================
Устанавливаем формат пикселей}
procedure TITLand.SetDCPixelFormat;
var
  nPixelFormat: Integer;
  pfd: TPixelFormatDescriptor;

begin
  FillChar(pfd, SizeOf(pfd), 0);

  pfd.dwFlags := PFD_DRAW_TO_WINDOW or PFD_SUPPORT_OPENGL or
                 PFD_DOUBLEBUFFER;

  nPixelFormat := ChoosePixelFormat(DC, @pfd);
  SetPixelFormat(DC, nPixelFormat, @pfd);
end;


procedure TITLand.ActiveFormCreate(Sender: TObject);
begin
  DC := GetDC(Handle);
  SetDCPixelFormat;
  hrc := wglCreateContext(DC);
  wglMakeCurrent(DC, hrc);
  Init;
  time := GetTickCount;
end;

procedure TITLand.ActiveFormDestroy(Sender: TObject);
begin
  glDeleteLists (Earth, 1);
  wglMakeCurrent(0, 0);
  wglDeleteContext(hrc);
  ReleaseDC(Handle, DC);
  DeleteDC (DC);
end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TITLand,
    Class_ITLand,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
