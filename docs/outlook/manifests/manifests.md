
# Outlook-Add-In-Manifeste
Erstellen Sie ein Manifest für ein Outlook-Add-In für Lese- oder Erstellformulare.

 _**Gilt für:** apps for Office | Office Add-ins | Outlook_

Ein Outlook-Add-In besteht aus zwei Komponenten: dem XML-Add-In-Manifest und einer Webseite, die durch die JavaScript-Bibliothek für Office-Add-Ins (office.js) unterstützt werden. Das Manifest beschreibt, wie das Add-In in Outlook-Clients integriert wird. Es sind nun drei Versionen des Manifestschemas verfügbar, einschließlich  **VersionOverrides**. Zu diesem Zeitpunkt sollten sich Entwickler auf das Erstellen eines Add-Ins mit Befehlen mithilfe von Manifestschemaversion 1.1 und  **VersionOverrides** 1.0 konzentrieren. Beispiel:

(../../images  Alle URL-Werte im folgenden Beispiel beginnen mit "YOUR_WEB_SERVER". Dieser Wert ist ein Platzhalter. In einem tatsächlichen gültige Manifest würden diese Werte gültige HTTPS-Webd-URLs enthalten.




```XML
<?xml version="1.0" encoding="UTF-8"?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft Outlook Dev Center</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Add-in Command Demo" />
  <Description DefaultValue="Adds command buttons to the ribbon in Outlook"/>
  <IconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-64.png" />
  <HighResolutionIconUrl DefaultValue="YOUR_WEB_SERVER/images/blue-80.png" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <!-- These elements support older clients that don't support add-in commands -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- NOTE: Just reusing the read taskpane page that is invoked by the button
             on the ribbon in clients that support add-in commands. You can 
             use a completely different page if desired -->
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="YOUR_WEB_SERVER/AppCompose/Home/Home.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">

        <DesktopFormFactor>
          <FunctionFile resid="functionFile" />
          
          <!-- Custom pane, only applies to read form -->
          <ExtensionPoint xsi:type="CustomPane">
            <RequestedHeight>100</RequestedHeight> 
            <SourceLocation resid="customPaneUrl"/>
            <Rule xsi:type="RuleCollection" Mode="Or">
              <Rule xsi:type="ItemIs" ItemType="Message"/>
              <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee"/>
            </Rule>
          </ExtensionPoint>
          
          <!-- Message compose form -->
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgComposeDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgComposeFunctionButton">
                  <Label resid="funcComposeButtonLabel" />
                  <Tooltip resid="funcComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="funcComposeSuperTipTitle" />
                    <Description resid="funcComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>addDefaultMsgToBody</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgComposeMenuButton">
                  <Label resid="menuComposeButtonLabel" />
                  <Tooltip resid="menuComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="menuComposeSuperTipTitle" />
                    <Description resid="menuComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgComposeMenuItem1">
                      <Label resid="menuItem1ComposeLabel" />
                      <Tooltip resid="menuItem1ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem1ComposeLabel" />
                        <Description resid="menuItem1ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg1ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgComposeMenuItem2">
                      <Label resid="menuItem2ComposeLabel" />
                      <Tooltip resid="menuItem2ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem2ComposeLabel" />
                        <Description resid="menuItem2ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg2ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgComposeMenuItem3">
                      <Label resid="menuItem3ComposeLabel" />
                      <Tooltip resid="menuItem3ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem3ComposeLabel" />
                        <Description resid="menuItem3ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg3ToBody</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                  <Label resid="paneComposeButtonLabel" />
                  <Tooltip resid="paneComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="paneComposeSuperTipTitle" />
                    <Description resid="paneComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="composeTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          
          <!-- Appointment compose form -->
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="apptComposeFunctionButton">
                  <Label resid="funcComposeButtonLabel" />
                  <Tooltip resid="funcComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="funcComposeSuperTipTitle" />
                    <Description resid="funcComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>addDefaultMsgToBody</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="apptComposeMenuButton">
                  <Label resid="menuComposeButtonLabel" />
                  <Tooltip resid="menuComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="menuComposeSuperTipTitle" />
                    <Description resid="menuComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="apptComposeMenuItem1">
                      <Label resid="menuItem1ComposeLabel" />
                      <Tooltip resid="menuItem1ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem1ComposeLabel" />
                        <Description resid="menuItem1ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg1ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="apptComposeMenuItem2">
                      <Label resid="menuItem2ComposeLabel" />
                      <Tooltip resid="menuItem2ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem2ComposeLabel" />
                        <Description resid="menuItem2ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg2ToBody</FunctionName>
                      </Action>
                    </Item>
                    <Item id="apptComposeMenuItem3">
                      <Label resid="menuItem3ComposeLabel" />
                      <Tooltip resid="menuItem3ComposeTip" />
                      <Supertip>
                        <Title resid="menuItem3ComposeLabel" />
                        <Description resid="menuItem3ComposeTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>addMsg3ToBody</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="apptComposeOpenPaneButton">
                  <Label resid="paneComposeButtonLabel" />
                  <Tooltip resid="paneComposeButtonTooltip" />
                  <Supertip>
                    <Title resid="paneComposeSuperTipTitle" />
                    <Description resid="paneComposeSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="composeTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          
          <!-- Message read form -->
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="msgReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Tooltip resid="funcReadButtonTooltip" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="msgReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Tooltip resid="menuReadButtonTooltip" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="msgReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Tooltip resid="menuItem1ReadTip" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Tooltip resid="menuItem2ReadTip" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="msgReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Tooltip resid="menuItem3ReadTip" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Tooltip resid="paneReadButtonTooltip" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          
          <!-- Appointment read form -->
          <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptReadDemoGroup">
                <Label resid="groupLabel" />
                <Tooltip resid="groupTooltip" />
                <!-- Function (UI-less) button -->
                <Control xsi:type="Button" id="apptReadFunctionButton">
                  <Label resid="funcReadButtonLabel" />
                  <Tooltip resid="funcReadButtonTooltip" />
                  <Supertip>
                    <Title resid="funcReadSuperTipTitle" />
                    <Description resid="funcReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="blue-icon-16" />
                    <bt:Image size="32" resid="blue-icon-32" />
                    <bt:Image size="80" resid="blue-icon-80" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getSubject</FunctionName>
                  </Action>
                </Control>
                <!-- Menu (dropdown) button -->
                <Control xsi:type="Menu" id="apptReadMenuButton">
                  <Label resid="menuReadButtonLabel" />
                  <Tooltip resid="menuReadButtonTooltip" />
                  <Supertip>
                    <Title resid="menuReadSuperTipTitle" />
                    <Description resid="menuReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="red-icon-16" />
                    <bt:Image size="32" resid="red-icon-32" />
                    <bt:Image size="80" resid="red-icon-80" />
                  </Icon>
                  <Items>
                    <Item id="apptReadMenuItem1">
                      <Label resid="menuItem1ReadLabel" />
                      <Tooltip resid="menuItem1ReadTip" />
                      <Supertip>
                        <Title resid="menuItem1ReadLabel" />
                        <Description resid="menuItem1ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemClass</FunctionName>
                      </Action>
                    </Item>
                    <Item id="apptReadMenuItem2">
                      <Label resid="menuItem2ReadLabel" />
                      <Tooltip resid="menuItem2ReadTip" />
                      <Supertip>
                        <Title resid="menuItem2ReadLabel" />
                        <Description resid="menuItem2ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getDateTimeCreated</FunctionName>
                      </Action>
                    </Item>
                    <Item id="apptReadMenuItem3">
                      <Label resid="menuItem3ReadLabel" />
                      <Tooltip resid="menuItem3ReadTip" />
                      <Supertip>
                        <Title resid="menuItem3ReadLabel" />
                        <Description resid="menuItem3ReadTip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="red-icon-16" />
                        <bt:Image size="32" resid="red-icon-32" />
                        <bt:Image size="80" resid="red-icon-80" />
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>getItemID</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
                <!-- Task pane button -->
                <Control xsi:type="Button" id="apptReadOpenPaneButton">
                  <Label resid="paneReadButtonLabel" />
                  <Tooltip resid="paneReadButtonTooltip" />
                  <Supertip>
                    <Title resid="paneReadSuperTipTitle" />
                    <Description resid="paneReadSuperTipDescription" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="green-icon-16" />
                    <bt:Image size="32" resid="green-icon-32" />
                    <bt:Image size="80" resid="green-icon-80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="readTaskPaneUrl" />
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
        <!-- Blue icon -->
        <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png"/>
        <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png"/>
        <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png"/>
        <!-- Red icon -->
        <bt:Image id="red-icon-16" DefaultValue="YOUR_WEB_SERVER/images/red-16.png"/>
        <bt:Image id="red-icon-32" DefaultValue="YOUR_WEB_SERVER/images/red-32.png"/>
        <bt:Image id="red-icon-80" DefaultValue="YOUR_WEB_SERVER/images/red-80.png"/>
        <!-- Green icon -->
        <bt:Image id="green-icon-16" DefaultValue="YOUR_WEB_SERVER/images/green-16.png"/>
        <bt:Image id="green-icon-32" DefaultValue="YOUR_WEB_SERVER/images/green-32.png"/>
        <bt:Image id="green-icon-80" DefaultValue="YOUR_WEB_SERVER/images/green-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
        <bt:Url id="readTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/TaskPane/TaskPane.html"/>
        <bt:Url id="composeTaskPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppCompose/TaskPane/TaskPane.html"/>
        <bt:Url id="customPaneUrl" DefaultValue="YOUR_WEB_SERVER/AppRead/CustomPane/CustomPane.html"/>"
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="Add-in Demo"/>
        <!-- Compose mode -->
        <bt:String id="funcComposeButtonLabel" DefaultValue="Insert default message"/>
        <bt:String id="menuComposeButtonLabel" DefaultValue="Insert message"/>
        <bt:String id="paneComposeButtonLabel" DefaultValue="Insert custom message"/>

        <bt:String id="funcComposeSuperTipTitle" DefaultValue="Inserts the default message"/>
        <bt:String id="menuComposeSuperTipTitle" DefaultValue="Choose a message to insert"/>
        <bt:String id="paneComposeSuperTipTitle" DefaultValue="Enter your own text to insert"/>
        
        <bt:String id="menuItem1ComposeLabel" DefaultValue="Hello World!"/>
        <bt:String id="menuItem2ComposeLabel" DefaultValue="Add-in commands are cool!"/>
        <bt:String id="menuItem3ComposeLabel" DefaultValue="Visit dev.outlook.com"/>

        <!-- Read mode -->
        <bt:String id="funcReadButtonLabel" DefaultValue="Get subject"/>
        <bt:String id="menuReadButtonLabel" DefaultValue="Get property"/>
        <bt:String id="paneReadButtonLabel" DefaultValue="Display all properties"/>

        <bt:String id="funcReadSuperTipTitle" DefaultValue="Gets the subject of the message or appointment"/>
        <bt:String id="menuReadSuperTipTitle" DefaultValue="Choose a property to get"/>
        <bt:String id="paneReadSuperTipTitle" DefaultValue="Get all properties"/>
        
        <bt:String id="menuItem1ReadLabel" DefaultValue="Get item class"/>
        <bt:String id="menuItem2ReadLabel" DefaultValue="Get date time created"/>
        <bt:String id="menuItem3ReadLabel" DefaultValue="Get item ID"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="groupTooltip" DefaultValue="Add-in Demo of the different command types"/>
        <!-- Compose mode -->
        <bt:String id="funcComposeButtonTooltip" DefaultValue="Inserts text into body of the message or appointment"/>
        <bt:String id="menuComposeButtonTooltip" DefaultValue="Inserts your choice of text into body of the message or appointment"/>
        <bt:String id="paneComposeButtonTooltip" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment"/>
        
        <bt:String id="funcComposeSuperTipDescription" DefaultValue="Inserts text into body of the message or appointment. This is an example of a function button."/>
        <bt:String id="menuComposeSuperTipDescription" DefaultValue="Inserts your choice of text into body of the message or appointment. This is an example of a drop-down menu button."/>
        <bt:String id="paneComposeSuperTipDescription" DefaultValue="Opens a pane where you can enter text to insert in the body of the message or appointment. This is an example of a button that opens a task pane."/>
        
        <bt:String id="menuItem1ComposeTip" DefaultValue="Inserts Hello World! into the body of the message or appointment." />
        <bt:String id="menuItem2ComposeTip" DefaultValue="Inserts Add-in commands are cool! into the body of the message or appointment." />
        <bt:String id="menuItem3ComposeTip" DefaultValue="Inserts Visit dev.outlook.com into the body of the message or appointment." />

        <!-- Read mode -->
        <bt:String id="funcReadButtonTooltip" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar"/>
        <bt:String id="menuReadButtonTooltip" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar"/>
        <bt:String id="paneReadButtonTooltip" DefaultValue="Opens a pane displaying all available properties of the message or appointment"/>
        
        <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment and displays it in the info bar. This is an example of a function button."/>
        <bt:String id="menuReadSuperTipDescription" DefaultValue="Gets the selected property of the message or appointment and displays it in the info bar. This is an example of a drop-down menu button."/>
        <bt:String id="paneReadSuperTipDescription" DefaultValue="Opens a pane displaying all available properties of the message or appointment. This is an example of a button that opens a task pane."/>
        
        <bt:String id="menuItem1ReadTip" DefaultValue="Gets the item class of the message or appointment and displays it in the info bar." />
        <bt:String id="menuItem2ReadTip" DefaultValue="Gets the date and time the message or appointment was created and displays it in the info bar." />
        <bt:String id="menuItem3ReadTip" DefaultValue="Gets the item ID of the message or appointment and displays it in the info bar." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>

```


## Schemaversionen

Nicht alle Outlook-Clients unterstützten sofort die neuesten Features, und einige Outlook-Benutzer verwenden eine ältere Version von Outlook. Dank Schemaversionen können Entwickler abwärtskompatible Add-Ins erstellen, die die neuesten Features verwenden, wo sie verfügbar sind, aber auch auf älteren Versionen funktionieren.

Das  **VersionOverrides**-Element in der Manifestdatei ist ein Beispiel dafür. Alle Elemente, die innerhalb von  **VersionOverrides** definiert sind, überschreiben dasselbe Element im anderen Teil des Manifests. Dies bedeutet, dass Outlook möglichst die Elemente aus dem **VersionOverrides**-Abschnitt verwendet, um das Add-In einzurichten. Wenn die Version von Outlook jedoch eine bestimmte Version von  **VersionOverrides** nicht unterstützt, ignoriert Outlook sie und verwendet die Informationen im Rest des Manifests.

Dieser Ansatz bedeutet, dass Entwickler nicht mehrere einzelne Manifeste erstellen müssen, sondern alles in einer Datei definieren können.

Die aktuellen Versionen des Schemas sind:


|||
|:-----|:-----|
|Version|Beschreibung|
|v1.0|Unterstützt Version 1.0 der JavaScript-API für Office. Für Outlook-Add-Ins wird das Leseformular unterstützt. |
|v1.1|Unterstützt Version 1.1 der JavaScript-API für Office und  **VersionOverrides**. Für Outlook-Add-Ins wird Unterstützung für das Entwurfsformular hinzugefügt.|
|**VersionOverrides** 1.0|Unterstützt neuere Versionen der JavaScript-API für Office. Dies unterstützt Add-In-Befehle.|
In diesem Artikel werden die Anforderungen für ein Manifest der Version 1.1 erläutert. Auch wenn Ihr Add-In-Manifest das  **VersionOverrides**-Element verwendet, ist es dennoch wichtig, Manifestelemente der Version 1.1 einzufügen, damit Ihre App mit älteren Clients funktioniert, die  **VersionOverrides** nicht unterstützten.


## Stammelement

Das Stammelement für das Outlook-Add-In-Manifest ist  **OfficeApp**. Dieses Element deklariert auch den Standard-Namespace, die Schemaversion und den Typ des Add-Ins. Platzieren Sie alle anderen Elemente im Manifest innerhalb der öffnenden und schließenden Tags. Der folgende Code ist ein Beispiel des Stammelements:


```XML
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- the rest of the manifest>

</OfficeApp>
```


## Version

Dies ist die Version des spezifischen Add-Ins. Wenn ein Entwickler etwas im Manifest aktualisiert, muss die Version ebenfalls erhöht werden. Auf diese Weise wird bei der Installation des neuen Manifests das vorhandene überschrieben, und der Benutzer erhält die neue Funktionalität. Wenn dieses Add-In an den Store gesendet wurde, muss das neue Manifest erneut gesendet und erneut überprüft werden. Dann erhalten Benutzer des Add-Ins das neue aktualisierte Manifest automatisch innerhalb weniger Stunden, nachdem es genehmigt wurde.

Wenn sich die für das Add-In benötigten Berechtigungen ändern, werden Benutzer aufgefordert, das Add-In zu aktualisieren und erneut die Zustimmung dafür zu geben. Wenn der Administrator dieses Add-In für die gesamte Organisation installiert hat, muss der Administrator zuerst erneut die Zustimmung geben. Benutzer verfügen inzwischen weiterhin über die alte Funktionalität.


## VersionOverrides

Das  **VersionOverrides** -Element, ist der Speicherort der Informationen für Add-In-Befehle. Weitere Informationen zu diesem Element finden Sie unter [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md).


## Lokalisierung

Manche Aspekte des Add-Ins müssen für unterschiedliche Gebietsschemas lokalisiert werden, z. B. der Name, die Beschreibung und die URL, die geladen wird. Diese Elemente können problemlos lokalisiert werden, indem Sie den Standardwert und dann Gebietsschemaüberschreibungen im  **Resources**-Element innerhalb des  **VersionOverrides**-Abschnitts festlegen. Im Folgenden wird gezeigt, wie Sie ein Bild, eine URL und eine Zeichenfolge überschreiben:


```XML
<Resources>
    <bt:Images>
      <bt:Image id="icon1_16x16" DefaultValue="https://contoso.com/images/app_icon_small.png" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/images/app_icon_small_arsa.png" />
        <!-- add information for other locales -->

    <bt:Urls>
      <bt:Url id="residDesktopFuncUrl" DefaultValue="https://contoso.com/urls/page_appcmdcode.html" >
        <bt:Override Locale="ar-sa" Value="https://contoso.com/urls/page_appcmdcode.html?lcid=ar-sa" />
        <!-- add information for other locales -->

    <bt:ShortStrings> 
      </bt:String>
      <bt:String id="residViewTemplates" DefaultValue="Launch My Add-in">
        <bt:Override Locale="ar-sa" Value="????? ???? ??????? ????????" />
        <!-- add information for other locales -->
    </bt:ShortStrings>

  </Resources>
```

Die Schemareferenz enthält umfassende Informationen dazu, welche Elemente lokalisiert werden können.


## Hosts

Outlook-Add-Ins geben das  **Hosts**-Element folgendermaßen an:


```XML
<OfficeApp>
...
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
...
</OfficeApp>
```

Dies unterscheidet sich von dem  **Hosts**-Element innerhalb des  **VersionOverrides**-Elements, das in [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md) erläutert wird.


## Anforderungen

Das  **Requirements**-Element gibt den Satz von APIs für das Add-In an. Für ein Outlook-Add-In muss der Anforderungssatz „Mailbox" und ein Wert von 1.1 oder höher angegeben werden. Die neueste Anforderungssatzversion finden Sie in der API-Referenz. Eine Erläuterung der Funktionsweise finden Sie unter [Outlook-Add-In-APIs](../outlook/apis.md).

Das  **Requirements**-Element kann auch im  **VersionOverrides**-Element angezeigt werden, sodass das Add-In eine andere Anforderung angeben kann, wenn es in Clients geladen wird, die  **VersionOverrides** unterstützen.

Im folgenden Beispiel wird das  **DefaultMinVersion**-Attribut des  **Sets**-Elements verwendet, um office.jx, Version 1.1 oder höher, anzufordern, und das  **MinVersion**-Attribut des  **Set**-Elements, um den Postfach-Anforderungssatz, Version 1.1, anzufordern.




```XML
<OfficeApp>
...
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="MailBox" MinVersion="1.1" />
    </Sets>
  </Requirements>
...
</OfficeApp>
```


## Formulareinstellungen

Das  **FormSettings**-Element wird von älteren Outlook-Clients verwendet, die nur Schema 1.1 und nicht  **VersionOverrides** unterstützen. Mithilfe dieses Elements definieren Entwickler, wie das Add-In in diesen Clients angezeigt wird. Es gibt zwei Teile – **ItemRead** und **ItemEdit**. Mit  **ItemRead** wird angegeben, wie das Add-In angezeigt wird, wenn der Benutzer Nachrichten und Termine liest. **ItemEdit** beschreibt, wie das Add-In angezeigt wird, während der Benutzer eine Antwort, eine neue Nachricht, einen neuen Termin erstellt oder einen Termin bearbeitet, dessen Organisator er ist.

Diese Einstellungen beziehen sich direkt auf die Aktivierungsregeln im  **Rule**-Element. Wenn ein Add-In beispielsweise angibt, dass es in einer gerade verfassten Nachricht angezeigt werden soll, muss ein  **ItemEdit**-Formular angegeben werden.

Weitere Informationen finden Sie in der [Schema reference for Office Add-ins manifests (v1.1)](http://msdn.microsoft.com/library/7e0cadc3-f613-8eb9-7ef-9032cbb97f92.aspx).


## App-Domänen

Die Domäne der Add-In-Startseite, die Sie im Element  **SourceLocation** angeben, ist die Standarddomäne. Wenn die Elemente **AppDomains** und **AppDomain** nicht verwendet werden und Ihr Add-In versucht, zu einer andere Domäne zu navigieren, öffnet der Browser ein neues Fenster außerhalb des Add-In-Bereichs. Wenn die Navigation innerhalb des Add-In-Bereichs bleiben soll, müssen Sie ein **AppDomains**-Element hinzufügen und jede zusätzliche Domäne in der Add-In-Manifestdatei in ein eigenes  **AppDomain**-Unterelement einbeziehen.

Das folgende Beispiel gibt eine Domäne  `https://www.contoso2.com` als zweite Domäne an, zu der das Add-In innerhalb des Add-In-Bereichs navigieren kann:




```XML
<OfficeApp>
...
  <AppDomains>
    <AppDomain>https://www.contoso2.com</AppDomain>
  </AppDomains>
...
</OfficeApp>
```

App-Domänen sind auch erforderlich, um die gemeinsame Nutzung von Cookies zwischen dem Popupfenster und dem im Rich-Client ausgeführten Add-In zu ermöglichen.


## Berechtigungen

Das  **Permissions**-Element enthält die erforderlichen Berechtigungen für das Add-In. In der Regel sollten Sie die erforderliche Mindestberechtigung angeben, die das Add-In benötigt, je nachdem, welche Methoden zum Verfassen Sie genau verwenden möchten. Für ein Mail-Add-In beispielsweise, das in einem Entwurfsformular aktiviert ist und Elementeigenschaften wie [item.requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) nur liest und nicht in diese schreibt, und [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md) nicht für den Zugriff auf Exchange Web Services-Vorgänge afruft, sollte die **ReadItem**-Berechtigung angegeben werden. Weitere Informationen zu den verfügbaren Berechtigungen finden Sie unter [Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](../../docs/outlook/privacy/understanding-outlook-add-in-permissions.md).


**Vierstufiges Berechtigungsmodell für Mail-Add-Ins**

![4-Stufen-Berechtigungsmodell für Mail-Apps-Schema v1.1](../../images/olowa15wecon_Permissions_4Tier.png)
```XML
<OfficeApp>
...
  <Permissions>ReadWriteItem</Permissions>
...
</OfficeApp>

```


## Aktivierungsregeln

Aktivierungsregeln werden im  **Rule**-Element angegeben. Das  **Rule**-Element kann als untergeordnetes Element des  **OfficeApp**-Elements in 1.1-Manifesten und zusätzlich als untergeordnetes Element des  **ExtensionPoint**-Elements in  **VersionOverrides** angezeigt werden. Weitere Informationen zur Verwendung dieses Elements in **VersionOverrides** finden Sie unter [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md).

Aktivierungsregeln können zum Aktivieren eines Add-Ins basierend auf einer oder mehrerer der folgenden Bedingungen in dem aktuell ausgewählten Element verwendet werden.


- Der Elementtyp und/oder die Nachrichtenklasse.
    
- Das Vorhandensein eines bestimmten Typs einer bekannten Entität, z. B. eine Adresse oder Telefonnummer
    
- Die Übereinstimmung eines regulären Ausdrucks im Nachrichtentext, Betreff oder der Absender-E-Mail-Adresse
    
- Das Vorhandensein einer Anlage
    
Weitere Informationen sowie Beispiele zu Aktivierungsregeln finden Sie unter [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md).


## Nächste Schritte: Add-In-Befehle


Nachdem Sie eine einfache Manifestdatei definiert haben, [definieren Sie Add-In-Befehle für Ihr Add-In](../outlook/manifests/define-add-in-commands.md). Add-In-Befehle werden als Schaltfläche im Menüband dargestellt, sodass Benutzer Ihr Add-In auf einfache und intuitive Weise aktivieren können. Weitere Informationen finden Sie unter [Add-In-Befehle für Outlook](../outlook/add-in-commands-for-outlook.md).


## Zusätzliche Ressourcen



- [Outlook-Add-Ins](../outlook/outlook-add-ins.md)
    
- [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md)
    
- [Lokalisierung für Office-Add-Ins](../../develop/localization.md)
    
- [Erstellen eines Mail-Add-ins für Outlook zur Ausführung auf Desktops, Tablets und mobilen Geräten (Schema v1.1)](8d425fb3-8a7c-429d-87b3-8046e964b153.md)
    
- [Datenschutz, Berechtigungen und Sicherheit für Outlook-Add-Ins](../../docs/develop/privacy-and-security.md)
    
- [Outlook-Add-In-APIs](../outlook/apis.md)
    
- [XML-Manifest für Office-Add-Ins](../../docs/overview/add-in-manifests.md)
    
- [Schemareferenz für Office-Add-in-Manifeste (v1. 1)](http://msdn.microsoft.com/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92%28Office.15%29.aspx)
    
- [Elementtypen und Meldungsklassen](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
    
- [Designrichtlinien für Office-Add-Ins](../add-in-design.md)
    
- [Angeben von Berechtigungen für den Outlook-Add-In-Zugriff auf die Benutzerpostfächer](../../docs/outlook/privacy/understanding-outlook-add-in-permissions.md)
    
- [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
