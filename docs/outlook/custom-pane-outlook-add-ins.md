
# Benutzerdefinierter Bereich Outlook-Add-Ins
Erfahren Sie, wie Sie ein Outlook-add-In definieren, das basierend auf bestimmten Regeln aktiviert wird.

 _**Gilt für:** apps for Office_

Ein benutzerdefinierter Bereich ist ein Erweiterungspunkt für ein Add-In, das aktiviert wird, wenn bestimmte Bedingungen in dem aktuell ausgewählten Element erfüllt sind. Er wird im Add-In-Manifest im  **VersionOverrides**-Element zusamen mit Add-In-Befehlen definiert, die vom Add-In implementiert werden. Weitere Informationen finden Sie unter [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md).
Ein benutzerdefinierter Bereich kann nur in Ansichten zum Lesen von Nachrichten oder Teilnehmeransichten von Terminen angezeigt werden. Es wird einen Eintrag in der Add-In-Leiste angezeigt. Wenn der Benutzer auf den Eintrag klickt, wird der benutzerdefinierte Bereich mit horizontaler Ausrichtung oberhalb des Nachrichtentexts des Elements angezeigt. Das Aussehen und Verhalten ist identisch mit den Lesemodus-Add-Ins, die keineAdd-In-Befehle implementieren.

**Ein Add-In mit einem benutzerdefinierten Bereich im Lesemodus**

![Shows a custom pane in a message read form.](../../images/c585ab0a-6c33-42d0-a20f-5deb8b54f480.png)Im folgenden Beispiel wird ein benutzerdefinierter Bereich für Elemente definiert, die Nachrichten sind oder eine Anlage haben oder eine Adresse enthalten. 



```
<ExtensionPoint xsi:type="CustomPane">
  <RequestedHeight>100< /RequestedHeight> 
  <SourceLocation resid="residReadTaskpaneUrl"/>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasAttachment"/>
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
  </Rule>
</ExtensionPoint>
```



-  **RequestedHeight** gibt die gewünschte Höhe dieses Mail-Add-Ins in Pixeln an, wenn dieses auf einem Desktopcomputer ausgeführt wird, und wird andernfalls ignoriert. Dieses Element kann einen Wert zwischen 32 und 450 haben. Wenn der Wert nicht festgelegt wird, beträgt der Standard 350 px. Optional.
    
-  **SourceLocation** gibt die HTML-Seite an, die die Benutzeroberfläche für den benutzerdefinierten Bereich bereitstellt. Das **resid**-Attribut wird auf den Wert des  **id**-Attributs eines  **Url**-Elements im  **Resources** -Element festgelegt. Erforderlich.
    
-  **Rule** gibt die Regel bzw. Regelsammlung an, die angibt, wann das Add-In aktiviert wird. Dasselbe Element wie in [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md) definiert, mit der Ausnahme, dass die [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx)-Regel die folgenden Änderungen aufweist:  **ItemType** ist entwender "Message" oder "AppointmentAttendee", und es gibt kein **FormType** -Attribut. Weitere Informationen finden Sie unter [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md).
    

## Zusätzliche Ressourcen



- [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Aktivierungsregeln für Outlook-Add-Ins](../outlook/manifests/activation-rules.md)
    
- [Outlook-Add-In-Manifeste](../outlook/manifests/manifests.md)
    
- [Definieren von Add-In-Befehlen in Ihrem Outlook-Add-In-Manifest](../outlook/manifests/define-add-in-commands.md)
    
