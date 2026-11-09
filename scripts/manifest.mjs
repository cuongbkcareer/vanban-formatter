// Sinh manifest.xml cho Word (Office Add-in, manifest XML 1.1: chạy trên Windows, Mac và Word web).
import { ADDIN_ID, REPO, SUPPORT_URL } from './config.mjs';

const esc = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Phiên bản manifest cần đúng dạng a.b.c.d */
const fourPart = (v) => {
  const p = v.split('.').map((n) => String(parseInt(n, 10) || 0));
  while (p.length < 4) p.push('0');
  return p.slice(0, 4).join('.');
};

export function generateManifest({ baseUrl, version }) {
  const base = baseUrl.replace(/\/+$/, '');
  if (!/^https:\/\//.test(base)) throw new Error('Địa chỉ host phải là HTTPS: ' + base);
  const origin = new URL(base).origin;
  const u = (p) => esc(`${base}/${p}`);

  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_ID}</Id>
  <Version>${fourPart(version)}</Version>
  <ProviderName>vanban-formatter</ProviderName>
  <DefaultLocale>vi-VN</DefaultLocale>
  <DisplayName DefaultValue="Chuẩn hoá văn bản hành chính"/>
  <Description DefaultValue="Chuyển văn bản thô thành văn bản hành chính đúng thể thức theo NĐ 30/2020/NĐ-CP. Xử lý ngay trên máy, không dùng AI, không gửi nội dung đi."/>
  <IconUrl DefaultValue="${u('assets/icon-32.png')}"/>
  <HighResolutionIconUrl DefaultValue="${u('assets/icon-80.png')}"/>
  <SupportUrl DefaultValue="${esc(SUPPORT_URL)}"/>
  <AppDomains>
    <AppDomain>${esc(origin)}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="WordApi" MinVersion="1.3"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${u('taskpane.html')}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="VanBanGroup">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="VanBanOpenTaskpane">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>VanBanTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${u('assets/icon-16.png')}"/>
        <bt:Image id="Icon.32" DefaultValue="${u('assets/icon-32.png')}"/>
        <bt:Image id="Icon.80" DefaultValue="${u('assets/icon-80.png')}"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${u('taskpane.html')}"/>
        <bt:Url id="Commands.Url" DefaultValue="${u('commands.html')}"/>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="${esc(`https://github.com/${REPO}#readme`)}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="Văn bản hành chính"/>
        <bt:String id="Button.Label" DefaultValue="Chuẩn hoá NĐ 30"/>
        <bt:String id="GetStarted.Title" DefaultValue="Chuẩn hoá văn bản hành chính"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Button.Tip" DefaultValue="Mở khung chuẩn hoá văn bản hành chính theo NĐ 30/2020/NĐ-CP"/>
        <bt:String id="GetStarted.Description" DefaultValue="Add-in đã sẵn sàng. Bấm nút Chuẩn hoá NĐ 30 ở thẻ Trang đầu (Home) để chuẩn hoá văn bản."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}
