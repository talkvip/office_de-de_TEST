
# Entwerfen von Office-Add-ins für das iPad
Stellen Sie Ihre Office-Add-Ins in Office für iPad zur Verfügung.

 _**Gilt für:** apps for Office | Office Add-ins | Office for iPad_

Office-Add-Ins stellen Benutzern im Kontext eines Office-Hosts zusätzliche Funktionen bereit. Um Ihre Office-Add-Ins in Office für iPad verfügbar zu machen, beachten Sie die folgenden Richtlinien:

- Stellen Sie sicher, dass das Add-In die nachfolgend aufgeführten Entwurfsrichtlinien erfüllt und den Vorschriften für Add-Ins im Office Store entspricht. Weitere Informationen finden Sie unter:
    
      - [Validierungsrichtlinien für an den Office Store übermittelte Apps und Add-Ins (Version 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx) (Richtlinien 3.4, 4.12, 10.8)
    
  - [Designrichtlinien für Office-Add-Ins](../add-in-design.md)
    
  - [Bewährte Designmethoden für Office Add-Ins](d455b76b-4d76-493d-a681-6b02ba1f38a8.md)
    
  - [Entwerfen für iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)
    
- Verwenden Sie Version 1.1 der JavaScript-API für Office. Weitere Informationen finden Sie unter [Aktualisieren der Version der JavaScript-API für Office und der Manifestschemadateien](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). Änderungen an der API werden unter [Änderungen in der JavaScript-API für Office](../../reference/what's-changed-in-the-javascript-api-for-office.md) beschrieben.
    
- Stellen Sie Ihr Office-Add-In kostenlos auf dem iPad zur Verfügung. Auf dem iPad sind nur kostenlose Office-Add-Ins zulässig. Sie können das gleiche Add-In auf anderen Plattformen verkaufen, wie z. B. Windows oder Office 365. Warum ein kostenloses Office-Add-In erstellen? Mit einem kostenlosen Add-In erreichen Sie mehr Benutzer.
    
- Schließen Sie keine kommerziellen Elemente in Ihr Add-In ein. Diese Richtlinie bedeutet, dass Sie Folgendes unterlassen müssen:
    
      - Anbieten von Testversionen.
    
  - Einschließen von UI-Elementen oder Links zu kostenpflichtigen Versionen des Add-Ins.
    
  - Einschließen von UI-Elementen oder Links zu Onlineshops, die zusätzliche Inhalte, Apps oder Add-Ins verkaufen.
    
  - Einschließen von UI-Elementen oder Speichern von Links in Ihren Seiten mit Datenschutzrichtlinie oder Nutzungsbedingungen.
    

     >**Hinweis**  Jeglicher Verstoß gegen diese Richtlinie führt dazu, dass Ihr Add-In sofort von Microsoft auf dem iPad deaktiviert wird.

    Sie können weitere Informationen zum Add-In oder Diensten innerhalb des Add-Ins teilen. Außerdem können Sie Ihre Add-Ins, die auf anderen Plattformen ausgeführt werden, kostenpflichtig machen oder kommerzielle Elemente darin aufnehmen. Zeigen Sie zu diesem Zweck verschiedene Benutzeroberflächen an, je nachdem, in welchem Browser oder auf welchem Gerät das Add-In ausgeführt wird. Verwenden Sie hierzu die folgenden Eigenschaften:
    
      - [Context.touchEnabled-Eigenschaft (JavaScript-API für Office)](http://msdn.microsoft.com/library/fd73f94b-7e4a-422c-afdb-fef6fba43766%28Office.15%29.aspx) – Erkennt, ob die Hostanwendung, in der das Add-In ausgeführt wird, Toucheingaben unterstützt.
    
  - [Context.commerceAllowed-Eigenschaft (JavaScript-API für Office)](http://msdn.microsoft.com/library/fd3812ac-14c3-485f-8991-d12fcc99c450%28Office.15%29.aspx) – Stellt fest, ob das Add-In auf einer Plattform ausgeführt wird, die kommerzielle Transaktionen einschränkt.
    
Wenn Sie sichergestellt haben, dass Ihr Add-In diese Anforderungen erfüllt, veröffentlichen Sie es erneut im Office Store. Weitere Informationen finden Sie unter [Veröffentlichen von Office- und SharePoint-Add-Ins und Office 365 Web Apps im Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx) und [Aktualisieren Ihrer Apps oder Add-Ins](http://msdn.microsoft.com/library/7313d32b-5345-4039-ac5d-a1ba0aef890b%28Office.15%29.aspx). Wählen Sie im Verkäuferdashboard die Option aus, mit der das Add-In auf dem iPad zur Verfügung gestellt wird, stimmen Sie der aktualisierten Store-Richtlinie zu, und geben Sie Ihre Apple-Entwickler-ID an. Lesen Sie außerdem die [Vereinbarung für Anbieter von Office Store-Anwendungen](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-us/Office_Store_Seller_Agreement_20120927.md). 

## Zusätzliche Ressourcen



- [Debuggen von Office-Add-Ins auf dem iPad und einem Mac-Computer](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    
