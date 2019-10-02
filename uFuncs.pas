unit uFuncs;

interface
uses
  windows,sysutils,strutils,classes,registry;
  procedure IEEmulator(VerCode: Integer);
  function IsWin64: Boolean;
implementation
{
10001 (0x2711)	Internet Explorer 10����ҳ��IE 10�ı�׼ģʽչ�֣�ҳ��!DOCTYPE��Ч
10000 (0x02710)	Internet Explorer 10����IE 10��׼ģʽ�а�����ҳ��!DOCTYPEָ������ʾ��ҳ��Internet Explorer 10 Ĭ��ֵ��
9999 (0x270F)	Windows Internet Explorer 9. ǿ��IE9��ʾ������!DOCTYPEָ��
9000 (0x2328)	Internet Explorer 9. Internet Explorer 9Ĭ��ֵ����IE9��׼ģʽ�а�����ҳ��!DOCTYPEָ������ʾ��ҳ��
8888 (0x22B8)	Internet Explorer 8��ǿ��IE8��׼ģʽ��ʾ������!DOCTYPEָ��
8000 (0x1F40)	Internet Explorer 8Ĭ�����ã���IE8��׼ģʽ�а�����ҳ��!DOCTYPEָ��չʾ��ҳ
7000 (0x1B58)	ʹ��WebBrowser Control�ؼ���Ӧ�ó�����ʹ�õ�Ĭ��ֵ����IE7��׼ģʽ�а�����ҳ��!DOCTYPEָ����չʾ��ҳ��

11001 (0x2AF9	Internet Explorer 11. Webpages are displayed in IE11 edge mode, regardless of the declared !DOCTYPE directive. Failing to declare a !DOCTYPE directive causes the page to load in Quirks.
11000 (0x2AF8)	IE11. Webpages containing standards-based !DOCTYPE directives are displayed in IE11 edge mode. Default value for IE11.
10001 (0x2711)	Internet Explorer 10. Webpages are displayed in IE10 Standards mode, regardless of the !DOCTYPE directive.
10000 (0x02710)	Internet Explorer 10. Webpages containing standards-based !DOCTYPE directives are displayed in IE10 Standards mode. Default value for Internet Explorer 10.
9999 (0x270F)	Windows Internet Explorer 9. Webpages are displayed in IE9 Standards mode, regardless of the declared !DOCTYPE directive. Failing to declare a !DOCTYPE directive causes the page to load in Quirks.
9000 (0x2328)	Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode. Default value for Internet Explorer 9.
Important  In Internet Explorer 10, Webpages containing standards-based !DOCTYPE directives are displayed in IE10 Standards mode.

8888 (0x22B8)	Webpages are displayed in IE8 Standards mode, regardless of the declared !DOCTYPE directive. Failing to declare a !DOCTYPE directive causes the page to load in Quirks.
8000 (0x1F40)	Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. Default value for Internet Explorer 8
Important  In Internet Explorer 10, Webpages containing standards-based !DOCTYPE directives are displayed in IE10 Standards mode.

7000 (0x1B58)	Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. Default value for applications hosting the WebBrowser Control.
}
procedure IEEmulator(VerCode: Integer);
const
  IE_SET_PATH_32='SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION';
  IE_SET_PATH_64='SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION';
var
  RegObj: TRegistry;
  sPath:string;
begin
  RegObj := TRegistry.Create;
  try
    //RegObj.RootKey := HKEY_CURRENT_USER;
    RegObj.RootKey := HKEY_LOCAL_MACHINE;
    RegObj.Access := KEY_ALL_ACCESS;
    if isWin64 then sPath := IE_SET_PATH_64 else sPath:=IE_SET_PATH_32;
    if not RegObj.OpenKey(sPath, False) then exit;
    try
      RegObj.WriteInteger(ExtractFileName(ParamStr(0)), VerCode);
    finally
      RegObj.CloseKey;
    end;
  finally
    RegObj.Free;
  end;
end;
{--}
{��Ҫע����GetNativeSystemInfo ������Windows XP ��ʼ���У�
 �� IsWow64Process ������ Windows XP with SP2 �Լ� Windows Server 2003 with SP1 ��ʼ���С�
 ����ʹ�øú�����ʱ�������GetProcAddress ��
}
function IsWin64: Boolean;
var
  Kernel32Handle: THandle;
  IsWow64Process: function(Handle: Windows.THandle; var Res: Windows.BOOL): Windows.BOOL; stdcall;
  GetNativeSystemInfo: procedure(var lpSystemInfo: TSystemInfo); stdcall;
  isWoW64: Bool;
  SystemInfo: TSystemInfo;
const
  PROCESSOR_ARCHITECTURE_AMD64 = 9;
  PROCESSOR_ARCHITECTURE_IA64 = 6;
begin
  Kernel32Handle := GetModuleHandle('KERNEL32.DLL');
  if Kernel32Handle = 0 then
    Kernel32Handle := LoadLibrary('KERNEL32.DLL');
  if Kernel32Handle <> 0 then
  begin
    IsWOW64Process := GetProcAddress(Kernel32Handle,'IsWow64Process');
    GetNativeSystemInfo := GetProcAddress(Kernel32Handle,'GetNativeSystemInfo');
    if Assigned(IsWow64Process) then
    begin
      IsWow64Process(GetCurrentProcess,isWoW64);
      Result := isWoW64 and Assigned(GetNativeSystemInfo);
      if Result then
      begin
        GetNativeSystemInfo(SystemInfo);
        Result := (SystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) or
                  (SystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64);
      end;
    end
    else Result := False;
  end
  else Result := False;
end;
end.
