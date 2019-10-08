unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.OleCtrls, SHDocVw,ActiveX,MSHTML,urlmon,
  strutils,
  Vcl.ComCtrls, Vcl.ExtCtrls;

type
  TfMain = class(TForm)
    Page1: TPageControl;
    Panel1: TPanel;
    btnLoad: TButton;
    Bar1: TStatusBar;
    ts0: TTabSheet;
    ts1: TTabSheet;
    ts2: TTabSheet;
    ts3: TTabSheet;
    Web1: TWebBrowser;
    Web2: TWebBrowser;
    Memo1: TMemo;
    Memo2: TMemo;
    edtUrl: TEdit;
    ts4: TTabSheet;
    Memo3: TMemo;
    btnCopyContent: TButton;
    btnMerge: TButton;
    ts5: TTabSheet;
    Memo4: TMemo;
    btnMerge2: TButton;
    ts6: TTabSheet;
    Web3: TWebBrowser;
    btnSaveas: TButton;
    Save1: TSaveDialog;
    本地内容: TTabSheet;
    Memo5: TMemo;
    web4: TWebBrowser;
    btnSave: TButton;
    ts8: TTabSheet;
    procedure btnLoadClick(Sender: TObject);
    procedure Web1DocumentComplete(ASender: TObject; const pDisp: IDispatch;
      const URL: OleVariant);
    procedure Web1NavigateComplete2(ASender: TObject; const pDisp: IDispatch;
      const URL: OleVariant);
    procedure Web2DocumentComplete(ASender: TObject; const pDisp: IDispatch;
      const URL: OleVariant);
    procedure btnCopyContentClick(Sender: TObject);
    procedure btnMergeClick(Sender: TObject);
    procedure btnMerge2Click(Sender: TObject);
    procedure btnSaveasClick(Sender: TObject);
    procedure Web3DocumentComplete(ASender: TObject; const pDisp: IDispatch;
      const URL: OleVariant);
    procedure btnSaveClick(Sender: TObject);
    procedure web4DocumentComplete(ASender: TObject; const pDisp: IDispatch;
      const URL: OleVariant);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Web1BeforeNavigate2(ASender: TObject; const pDisp: IDispatch;
      const URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
      var Cancel: WordBool);
  private
    { Private declarations }
    procedure AppException(Sender: TObject; E: Exception);

  public
    { Public declarations }
  end;
const
  SOFT_TITLE='网页修改保存工具';
  SOFT_VERSION='v1.0';
var
  fMain: TfMain;

  mProtocol,mSite,mTitle,mCharset,mPage,mWorkDir,mModelPage:string;
function getPageCode(doc:IHTMLDocument2):tstrings;//返回页面源代码
function getPageContent(doc:IHTMLDocument2;tag:string):tstrings;//返回页面源代码
function DownloadToFile(Source, Dest: string): Boolean;
procedure SaveAsPage(doc:IHTMLDocument2;localPageName:string);//页面另存为
procedure SaveSinglePage(doc:IHTMLDocument2;localPageName:string);//仅仅保存一个页面
procedure  WB_LoadHTML(WebBrowser:  TWebBrowser;  HTMLCode:  string);//将 HTML 代码直接加入到 TWebbrowser 组件中去
procedure  WBLoadHTML(WebBrowser:  TWebBrowser;  HTMLCode:  tstrings);
implementation

{$R *.dfm}
uses
  uFuncs;
procedure TfMain.AppException(Sender: TObject; E: Exception);
begin

  //Log(e.Message);
end;
//从内存中加载页面（比加载htm文件速度快）uses ActiveX;
procedure  WBLoadHTML(WebBrowser:  TWebBrowser;  HTMLCode:  tstrings);
var
   ms:  TMemoryStream;
begin
   if  not Assigned(WebBrowser.Document)  then
      WebBrowser.Navigate('about:blank');
   if  Assigned(WebBrowser.Document)  then
   begin
       try
           ms  :=  TMemoryStream.Create;
           try
               HTMLCode.SaveToStream(ms,tEncoding.UTF8);
               ms.Seek(0,  0);
               (WebBrowser.Document  as  IPersistStreamInit).Load(TStreamAdapter.Create(ms));
               finally
               ms.Free;
           end;
       finally
       end;
   end;
end;

function getPageContent(doc:IHTMLDocument2;tag:string):tstrings;//返回页面源代码
var
  ms: TMemoryStream;
  ss:tstrings;
  content:ihtmlelement;
begin
  ss:=tstringlist.Create;
  content:=(doc.all.item(tag,0) as ihtmlelement);
  content.style.removeAttribute('HEIGHT',0);
  ss.Text:=content.outerHTML;
  result:=ss;
end;
function getPageCode(doc:IHTMLDocument2):tstrings;//返回页面源代码
var
  ms: TMemoryStream;
  ss:tstrings;
begin
 ms := TMemoryStream.Create;
 ss:=tstringlist.Create;
 (doc as IPersistStreamInit).Save(TStreamAdapter.Create(ms), True);
 ms.Position := 0;
 ss.LoadFromStream(ms,TEncoding.UTF8);
 ms.Free;
 result:=ss;
end;

procedure TfMain.btnCopyContentClick(Sender: TObject);
var
  doc:IHTMLDocument2;
begin
  doc:=web1.Document as IHTMLDocument2; //得到接口；
  memo3.Lines:=getpagecontent(doc,'article_content');
  bar1.Panels[0].Text:='加载完毕！';
  page1.ActivePageIndex:=4;
end;

procedure TfMain.btnLoadClick(Sender: TObject);
begin
  page1.ActivePageIndex:=0;
  web1.Navigate(edturl.Text);
  bar1.Panels[0].Text:='正在加载页面...';
  web2.Navigate(mModelPage);
  //web2.Navigate(memo2.Lines.Text);
end;

procedure TfMain.btnMerge2Click(Sender: TObject);
var
  mergefilename:string;
begin
  memo4.Lines:=memo2.Lines;
  memo4.Lines.Insert(12,memo3.Text);
  mergefilename:=mWorkDir+'\tmp.htm';
  memo4.Lines.SaveToFile(mergefilename,tEncoding.UTF8);
  bar1.Panels[0].Text:='正在加载合成页面...';
  page1.ActivePageIndex:=6;
  //web3.Navigate(mergefilename);
  wbloadhtml(web3,memo4.Lines);
end;

procedure TfMain.btnMergeClick(Sender: TObject);
var
  doc1,doc2:IHTMLDocument2;
  ss:tstrings;
begin
  doc1:=web1.Document as IHTMLDocument2; //得到接口；
  ss:=getpagecontent(doc1,'article_content');
  memo1.Lines:=ss;
  doc2:=web2.Document as IHTMLDocument2; //得到接口；
  doc2.body.innerHTML:=ss.Text;
  //doc2.body.innerHTML:='aa';

  //memo4.Lines.Text:=doc2.all.toString;
  memo4.Lines:=getpagecode(doc2);
end;

procedure TfMain.btnSaveasClick(Sender: TObject);
var
  filename:string;
  doc3:IHTMLDocument2;
begin
  //DownloadToFile('https://stackedit.io/style.css','C:\tmp\2\style.css');
  if not Assigned(web3.Document) then Exit;
  if(save1.Execute())then begin
    filename:=save1.FileName;
    save1.InitialDir:=extractfiledir(filename);
    doc3:=web3.Document as IHTMLDocument2;
    SaveAsPage(doc3,filename);
    fmain.memo5.Lines.Text:=fmain.web3.OleObject.Document.all.tags('HTML').Item(0).outerHTML;
    memo5.Lines.SaveToFile(filename,tEncoding.UTF8);
    page1.ActivePageIndex:=8;
    web4.Navigate(filename);
    mPage:=filename;
  end;
end;

procedure TfMain.btnSaveClick(Sender: TObject);
begin
  page1.ActivePageIndex:=8;
  bar1.Panels[0].Text:='正在加载本地页面...！';
  memo5.Lines.SaveToFile(mPage,tEncoding.UTF8);
  web4.Navigate(mPage);
end;

procedure TfMain.FormCreate(Sender: TObject);
begin
  Application.OnException := AppException;
  //Set8087CW(Longword($133f));
  IEEmulator(11001);
end;

procedure TfMain.FormShow(Sender: TObject);
begin
  web2.Navigate(mModelPage);
  fmain.caption:=SOFT_TITLE+SOFT_VERSION;
end;

procedure TfMain.Web1BeforeNavigate2(ASender: TObject; const pDisp: IDispatch;
  const URL, Flags, TargetFrameName, PostData, Headers: OleVariant;
  var Cancel: WordBool);
begin
  bar1.Panels[0].Text:='正在加载页面...';
end;

procedure TfMain.Web1DocumentComplete(ASender: TObject; const pDisp: IDispatch;
  const URL: OleVariant);
var
  doc:IHTMLDocument2;
begin
  if(Web1.ReadyState<>READYSTATE_COMPLETE)then exit;
  doc:=web1.Document as IHTMLDocument2; //得到接口；
  memo1.Lines:=getpagecode(doc);
  mProtocol:=doc.protocol;
  mTitle:=doc.title;
  mSite:=doc.domain;
  mCharset:=doc.charset;
  if(mProtocol='HyperText Transfer Protocol with Privacy')then mProtocol:='https://' else mProtocol:='http://';
  fmain.caption:=SOFT_TITLE+SOFT_VERSION+'('+mTitle+')';
  bar1.Panels[0].Text:='远程页面加载完毕！';
  //page1.ActivePageIndex:=1;
end;

procedure TfMain.Web1NavigateComplete2(ASender: TObject; const pDisp: IDispatch;
  const URL: OleVariant);
begin
  web1.Silent := True;
end;

procedure TfMain.Web2DocumentComplete(ASender: TObject; const pDisp: IDispatch;
  const URL: OleVariant);
var
  doc:IHTMLDocument2;
begin
  doc:=web2.Document as IHTMLDocument2; //得到接口；
  memo2.Lines:=getpagecode(doc);
  bar1.Panels[0].Text:='模板页面加载完毕！';

end;
procedure TfMain.Web3DocumentComplete(ASender: TObject; const pDisp: IDispatch;
  const URL: OleVariant);
var
  doc:IHTMLDocument2;
begin
  doc:=web3.Document as IHTMLDocument2; //得到接口；
  doc.title:=mTitle;
  bar1.Panels[0].Text:='合成页面加载完毕！';
end;

procedure TfMain.web4DocumentComplete(ASender: TObject; const pDisp: IDispatch;
  const URL: OleVariant);
begin
   bar1.Panels[0].Text:='本地页面加载完毕！';
end;

//------------------------------------------页面另存为区------------------------------------------
procedure SaveAsPage(doc:IHTMLDocument2;localPageName:string);//页面另存为
Var
  all:IHTMLElementCollection;
  sheets:IHTMLstyleSheetsCollection;
  len,I,p:integer;
  item:OleVariant;
  url,newUrl,filename,localfilename,localPageDir,localImageDir,filetag,fileext,num:string;
  ss:tstrings;
begin
  //网页中的图片文件：
  localPageDir:=extractfiledir(localPageName);
  localfilename:=extractfilename(localPageName);
  filetag:=leftstr(localfilename,length(localfilename)-4);
  localImageDir:=localPageDir+'\images';
  if(not directoryexists(localImageDir))then forcedirectories(localImageDir);
  all:=doc.images;
  len:=all.length;
  for I:=0 to len-1 do begin
    item:=all.item(I,varempty);
    url:=item.src;
    url:=trim(url);
    if(length(url)=0)then continue;
    if(pos('/',url)=1)then url:=mProtocol+mSite+url;
    //replacestr(url,'file://',mProtocol);
    if(leftstr(url,7)='file://')then url:=replacestr(url,'file://',mProtocol);
    num:=inttostr(i+1);
    if(length(num)=1)then num:='0'+num;
    if(url[length(url)-3]='.')then fileext:=rightstr(url,4) else fileext:='.jpg';
    filename:=filetag+num+fileext;
    localfilename:=localImageDir+'\'+filename;
    DownloadToFile(url,localfilename);
    newUrl:='images/'+filename;
    item.src:=newUrl;
  end;
  //ss:=tstringlist.Create;
  //ss.Text:=doc.body.outerHTML;
  //ss.SaveToFile(localpagename,TEncoding.UTF8);
  //ss.Free;
  //SaveSinglePage(doc,localpagename);
  //fmain.memo4.Lines:=getPagecode(doc);
end;
procedure SaveSinglePage(doc:IHTMLDocument2;localPageName:string);//仅仅保存一个页面
var
  ms: TMemoryStream;
  ss:tstrings;
begin
 ms := TMemoryStream.Create;
 ss:=tstringlist.Create;
 (doc as IPersistStreamInit).Save(TStreamAdapter.Create(ms), True);
 ms.Position := 0;
 ss.LoadFromStream(ms,TEncoding.UTF8);
 ss.SaveToFile(localPageName,TEncoding.UTF8);
 ms.Free;
 ss.Free;
end;
//------------------------------------------公共函数区----------------------------------------------
//uses urlmon;
function DownloadToFile(Source, Dest: string): Boolean;
begin
  try
    Result := UrlDownloadToFile(nil, PChar(source), PChar(Dest), 0, nil) = 0;
  except
    Result := False;
  end;
end;
procedure  WB_LoadHTML(WebBrowser:  TWebBrowser;  HTMLCode:  string);

var

   sl:  TStringList;

   ms:  TMemoryStream;

begin

   WebBrowser.Navigate('about:blank');

   if  Assigned(WebBrowser.Document)  then

   begin

       sl  :=  TStringList.Create;

       try

           ms  :=  TMemoryStream.Create;

           try

               sl.Text  :=  HTMLCode;

               sl.SaveToStream(ms);

               ms.Seek(0,  0);

               (WebBrowser.Document  as  IPersistStreamInit).Load(TStreamAdapter.Create(ms));

              finally

               ms.Free;

           end;

       finally

           sl.Free;

       end;

   end;

end;
procedure init();
var
  filename:string;
begin
  filename:=application.ExeName;
  mWorkDir:=extractfiledir(filename);
  mModelPage:=mWorkDir+'\csdnmodel.htm';
end;
initialization
  OleInitialize(nil);
  init();
finalization
  OleUninitialize;
end.
