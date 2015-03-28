unit ItlandWImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, ItlandWProj1_TLB, StdVcl, ExtCtrls, OpenGL;

type
  TItlandW = class(TActiveForm, IItlandW)
    Timer1: TTimer;
    procedure ActiveFormCreate(Sender: TObject);
    procedure ActiveFormDestroy(Sender: TObject);
    procedure ActiveFormPaint(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    FEvents: IItlandWEvents;
    procedure ActivateEvent(Sender: TObject);
    procedure ClickEvent(Sender: TObject);
    procedure CreateEvent(Sender: TObject);
    procedure DblClickEvent(Sender: TObject);
    procedure DeactivateEvent(Sender: TObject);
    procedure DestroyEvent(Sender: TObject);
    procedure KeyPressEvent(Sender: TObject; var Key: Char);
    procedure PaintEvent(Sender: TObject);
  protected
    { Protected declarations }
    procedure WMSize(var Msg: TWMPaint); message WM_SIZE;
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
Var
  Message : TMsg;
  DC : HDC;
  hrc : HGLRC;
  MyPaint : TPaintStruct;
  Angle : GLint = 0;

const
  Sphere = 1;

implementation

uses ComObj, ComServ;

{$R *.DFM}

{ TItlandW }

procedure TItlandW.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_ItlandWPage); }
end;

procedure TItlandW.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as IItlandWEvents;
  inherited EventSinkChanged(EventSink);
end;

procedure TItlandW.Initialize;
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

function TItlandW.Get_Active: WordBool;
begin
  Result := Active;
end;

function TItlandW.Get_AlignDisabled: WordBool;
begin
  Result := AlignDisabled;
end;

function TItlandW.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TItlandW.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TItlandW.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TItlandW.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TItlandW.Get_Color: OLE_COLOR;
begin
  Result := OLE_COLOR(Color);
end;

function TItlandW.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TItlandW.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TItlandW.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TItlandW.Get_Font: IFontDisp;
begin
  GetOleFont(Font, Result);
end;

function TItlandW.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TItlandW.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TItlandW.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TItlandW.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TItlandW.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TItlandW.Get_ScreenSnap: WordBool;
begin
  Result := ScreenSnap;
end;

function TItlandW.Get_SnapBuffer: Integer;
begin
  Result := SnapBuffer;
end;

function TItlandW.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TItlandW.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TItlandW._Set_Font(var Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TItlandW.ActivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnActivate;
end;

procedure TItlandW.ClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnClick;
end;

procedure TItlandW.CreateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnCreate;
end;

procedure TItlandW.DblClickEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDblClick;
end;

procedure TItlandW.DeactivateEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDeactivate;
end;

procedure TItlandW.DestroyEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnDestroy;
end;

procedure TItlandW.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  TempKey := Smallint(Key);
  if FEvents <> nil then FEvents.OnKeyPress(TempKey);
  Key := Char(TempKey);
end;

procedure TItlandW.PaintEvent(Sender: TObject);
begin
  if FEvents <> nil then FEvents.OnPaint;
end;

procedure TItlandW.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TItlandW.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TItlandW.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TItlandW.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TItlandW.Set_Color(Value: OLE_COLOR);
begin
  Color := TColor(Value);
end;

procedure TItlandW.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TItlandW.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TItlandW.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TItlandW.Set_Font(const Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TItlandW.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TItlandW.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TItlandW.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TItlandW.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TItlandW.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TItlandW.Set_ScreenSnap(Value: WordBool);
begin
  ScreenSnap := Value;
end;

procedure TItlandW.Set_SnapBuffer(Value: Integer);
begin
  SnapBuffer := Value;
end;

procedure TItlandW.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

procedure TItlandW.ActiveFormCreate(Sender: TObject);
 //**************************************
procedure SetDCPixelFormat (hdc : HDC);
var
 pfd : TPixelFormatDescriptor; // данные формата пикселей
 nPixelFormat : Integer;
Begin
 FillChar(pfd, SizeOf(pfd), 0);

 pfd.dwFlags := PFD_DRAW_TO_WINDOW or PFD_SUPPORT_OPENGL or
                 PFD_DOUBLEBUFFER;
 nPixelFormat := ChoosePixelFormat (hdc, @pfd); // запрос системе - поддерживается ли выбранный формат пикселей
 SetPixelFormat (hdc, nPixelFormat, @pfd);      // устанавливаем формат пикселей в контексте устройства
End;
  //*************************************
function ReadBitmap(const FileName : String;
                    var sWidth, tHeight: GLsizei): pointer;
const
  szh = SizeOf(TBitmapFileHeader);
  szi = SizeOf(TBitmapInfoHeader);
type
  TRGB = record
    r, g, b : GLbyte;
  end;
  TWrap = Array [0..0] of TRGB;
var
  BmpFile : File;
  bfh : TBitmapFileHeader;
  bmi : TBitmapInfoHeader;
  x, size: GLint;
  temp: GLbyte;
begin
  AssignFile (BmpFile, FileName);
  Reset (BmpFile, 1);
  Size := FileSize (BmpFile) - szh - szi;
  Blockread(BmpFile, bfh, szh);
  BlockRead (BmpFile, bmi, szi);
  If Bfh.bfType <> $4D42 then begin
    MessageBox(Handle, 'Invalid Bitmap', 'Error', MB_OK);
    Result := nil;
    Exit;
  end;
  sWidth := bmi.biWidth;
  tHeight := bmi.biHeight;
  GetMem (Result, Size);
  BlockRead(BmpFile, Result^, Size);
  For x := 0 to sWidth*tHeight-1 do
    With TWrap(result^)[x] do begin
      temp := r;
      r := b;
      b := temp;
  end;
end;
 //**************************************
procedure Init;
const
 LightPos : Array [0..3] of GLfloat = (10.0, 10.0, 10.0, 1.0);
var
 Quadric : GLUquadricObj;
 wrkPointer : Pointer;
 sWidth, tHeight : GLsizei;
begin
 glEnable(GL_LIGHTING);
 glEnable(GL_LIGHT0);
 glLightfv(GL_LIGHT0, GL_POSITION, @LightPos);

 Quadric := gluNewQuadric;
 gluQuadricTexture (Quadric, TRUE);

 glNewList (Sphere, GL_COMPILE);
   gluSphere (Quadric, 1.0, 24, 12);
 glEndList;

 gluDeleteQuadric (Quadric);

 wrkPointer := ReadBitmap('C:\earth.bmp', sWidth, tHeight);

 glTexImage2D(GL_TEXTURE_2D, 0, 3, sWidth, tHeight, 0,
              GL_RGB, GL_UNSIGNED_BYTE, wrkPointer);

 Freemem(wrkPointer);

 glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST);
 glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST);
 glEnable(GL_TEXTURE_2D);

 glEnable(GL_DEPTH_TEST);
end;
begin
  DC := GetDC (Handle);
       SetDCPixelFormat (DC);        // установить формат пикселей
       hrc := wglCreateContext (DC); // создаёт контекст воспроизведения OpenGL
       wglMakeCurrent (DC, hrc);     // устанавливает текущий контекст воспроизведения
       Init;
end;

procedure TItlandW.ActiveFormDestroy(Sender: TObject);
begin
       glDeleteLists (Sphere, 1);
       wglMakeCurrent (0, 0);
       wglDeleteContext (hrc);
       ReleaseDC (Handle, DC);
       DeleteDC (DC);
       PostQuitMessage (0);
       Exit;
end;

procedure TItlandW.ActiveFormPaint(Sender: TObject);
begin
    DC := BeginPaint (Handle, MyPaint);

       glClear( GL_COLOR_BUFFER_BIT or GL_DEPTH_BUFFER_BIT );
       glPushMatrix;
         glRotatef (Angle, 0.0, 0.0, 1.0);
         glCallList(1);
       glPopMatrix;

       SwapBuffers(DC);
       EndPaint (Handle, MyPaint);
end;
procedure TItlandW.WMSize(var Msg: TWMPaint);
begin
      glViewport(0, 0, ClientWidth, ClientHeight );
       glMatrixMode( GL_PROJECTION );
       glLoadIdentity;
       glFrustum( -1.0, 1.0, -1.0, 1.0, 5.0, 15.0 );
       glMatrixMode( GL_MODELVIEW );
       glLoadIdentity;
       glTranslatef( 0.0, 0.0, -12.0 );
       glRotatef(-90.0, 1.0, 0.0, 0.0);
end;

procedure TItlandW.Timer1Timer(Sender: TObject);
begin
Angle := (Angle + 1) mod 360;
  InvalidateRect(Handle, nil, False);
end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TItlandW,
    Class_ItlandW,
    1,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
