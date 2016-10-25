
# Membros de Workbook (Excel)


Representa uma pasta de trabalho do Microsoft Excel.


## Eventos



|**Nome**|**Descrição**|
|:-----|:-----|
|[Activate](74bb6d8c-aec8-7bb6-5c30-9a20f9a7afe8.md)|Ocorre quando uma pasta de trabalho, planilha, planilha de gráfico ou gráfico inserido é ativado.|
|[AddinInstall](671117b2-590e-9d6f-29ae-5f0bf30d4e99.md)|Ocorre quando a pasta de trabalho é instalada como um suplemento|
|[AddinUninstall](e35ba67b-3e04-d950-2f8b-141e478ddb67.md)|Ocorre quando a pasta de trabalho é desinstalada como um suplemento.|
|[AfterSave](97fee36a-f77c-29ab-de1d-b6069b2d74d8.md)|Ocorre depois de a pasta de trabalho ser salva.|
|[AfterXmlExport](fe1e0a53-9f4e-ac88-58f7-fe420e57cabd.md)|Ocorre após o Microsoft Excel salvar ou exportar dados XML da pasta de trabalho especificada.|
|[AfterXmlImport](b43adf53-6b67-6127-e69d-6ea05f68b7f6.md)|Ocorre após a atualização de uma conexão de dados XML existente ou após a importação de novos dados XML em uma pasta de trabalho especificada do Microsoft Excel.|
|[BeforeClose](1c440637-8289-c6dd-24e0-1b2764fd1694.md)|Ocorre antes da pasta de trabalho fechar. Se a pasta de trabalho tiver sido alterada, este evento ocorrerá antes do usuários ser solicitado a salvar as alterações.|
|[BeforePrint](2c97cb32-2bb3-2848-b5ed-32d9129af080.md)|Ocorre antes da pasta de trabalho (ou qualquer coisa que ela contenha) ser impressa.|
|[BeforeSave](dfa3e20f-1fb2-f84f-4b92-a98f22b6e637.md)|Ocorre antes da pasta de trabalho ser salva.|
|[BeforeXmlExport](ee2af5de-e52f-9434-aa7c-5dc9bb102d1b.md)|Ocorre antes de o Microsoft Excel salvar ou exportar dados XML da pasta de trabalho especificada.|
|[BeforeXmlImport](a0a589c6-15f9-5599-c0b6-c6f881816ad6.md)|Ocorre antes da atualização de uma conexão de dados XML existente ou antes dos novos dados XML serem importados em uma pasta de trabalho do Microsoft Excel.|
|[Deactivate](6bd5411c-ac43-95cf-6755-49780ac765e9.md)|Ocorre quando a planilha, a pasta de trabalho ou o gráfico é desativado.|
|[ModelChange](efe01088-273b-f9d8-ea3e-2ea1725ba7b2.md)|Ocorre depois que o modelo de dados do Excel é alterado.|
|[NewChart](76e7f325-9244-fd8c-b38d-063f0193a5e9.md)|Ocorre quando um novo gráfico é criado na pasta de trabalho.|
|[NewSheet](5abb254d-a2c3-7dac-e79f-0de74a081ecd.md)|Ocorre quando uma nova planilha é criada na pasta de trabalho.|
|[Open](313adc5e-0319-4ca4-cf5d-791b7184dacf.md)|Ocorre quando a pasta de trabalho é aberta.|
|[PivotTableCloseConnection](e267ab5b-382e-b270-18c8-f643e03e4604.md)|Ocorre após um relatório de tabela dinâmica fechar a conexão com sua fonte de dados.|
|[PivotTableOpenConnection](b6ce12f7-7bc6-bfcc-33f4-2e8ea6e53bae.md)|Ocorre após um relatório de tabela dinâmica abrir a conexão com sua fonte de dados.|
|[RowsetComplete](05bdddba-6716-4bba-01b6-863f27623821.md)|O evento é acionado quando o usuário analisa o conjunto de registros ou invoca a ação do conjunto de linhas em uma tabela dinâmica OLAP.|
|[SheetActivate](2a7c05c3-5b66-8012-5ac5-981dcfc7f947.md)|Ocorre quando uma planilha é ativada.|
|[SheetBeforeDelete](42406738-0fcd-4ef7-9bd6-abcc05f5e922.md)||
|[SheetBeforeDoubleClick](69d21025-78ef-deab-39be-b7a092d611f5.md)|Ocorre quando uma planilha é clicada duas vezes, antes da ação padrão de clique duplo.|
|[SheetBeforeRightClick](d84dd9fd-85d3-009e-281b-cfc0d2874859.md)|Ocorre quando uma planilha é clicada com o botão direito, antes da ação padrão de clique com o botão direito.|
|[SheetCalculate](0610bfa5-15dc-a57f-f362-cf897bd54b91.md)|Ocorre depois de uma planilha ser recalculada ou depois de os dados alterados serem plotados em um gráfico.|
|[SheetChange](37e727d8-255c-ac23-45d8-13a8e7639991.md)|Ocorre quando células de uma planilha são alteradas pelo usuário ou por um link externo.|
|[SheetDeactivate](befde22b-69ce-c34f-2b9e-da5e026972e3.md)|Ocorre quando uma planilha é desativada.|
|[SheetFollowHyperlink](be29df8c-4e8e-f719-ae1d-f91a11b89491.md)|Ocorre quando você clica em qualquer hiperlink no Microsoft Excel. Para eventos de nível de planilha, consulte o tópico de ajuda para o evento  **[FollowHyperlink](c63eec19-008e-bfb5-1357-3d02426c1bab.md)**.|
|[SheetLensGalleryRenderComplete](8ac48e9f-7a15-c674-6d96-e9c1466473bc.md)|Ocorre quando os ícones de uma galeria texto explicativo (dinâmicos e estáticos) concluiu a renderização para uma planilha.|
|[SheetPivotTableAfterValueChange](8460f5f1-d415-7aac-6a3d-fa0944036e9c.md)|Ocorre depois que uma célula ou intervalo de células dentro de uma Tabela Dinâmica é editado ou recalculado (para células que contêm fórmulas).|
|[SheetPivotTableBeforeAllocateChanges](2f767b5b-27fb-33de-c91d-76bbc52ea171.md)|Ocorre antes que as alterações sejam aplicadas a uma Tabela Dinâmica.|
|[SheetPivotTableBeforeCommitChanges](7e189a4f-a349-f862-375a-fa66311629cc.md)|Ocorre antes de as alterações serem confirmadas com base na fonte de dados OLAP de uma Tabela Dinâmica.|
|[SheetPivotTableBeforeDiscardChanges](e8f1ae21-c9ed-6f4d-a85c-d6768060a66f.md)|Ocorre antes de as alterações feitas em uma Tabela Dinâmica serem descartadas.|
|[SheetPivotTableChangeSync](c280b935-3dbf-0666-b727-64d6b4ac7ebd.md)|Ocorre depois que são feitas alterações em uma Tabela Dinâmica.|
|[SheetPivotTableUpdate](0b37939a-28dd-ef8b-ea5e-fc3768f8979a.md)|Ocorre depois que a planilha do relatório de tabela dinâmica tiver sido atualizada.|
|[SheetSelectionChange](a3829af1-2917-9526-1d64-91eeb6c198ce.md)|Ocorre quando a seleção é alterada em uma planilha (não ocorre caso a seleção esteja em uma planilha de gráfico).|
|[SheetTableUpdate](609d331e-45b9-885b-a395-d80ccf4c19a5.md)|Ocorre depois que a tabela de planilha foi atualizada.|
|[Sync](ce8b77e1-a316-c0e3-f0f8-ce4ac22ec430.md)|Este objeto ou membro foi substituído, mas continua a fazer parte do modelo de objeto para compatibilidade com versões anteriores. Você não deve usá-lo nos novos aplicativos.|
|[WindowActivate](e99d955c-1975-44c3-05b3-3aa6e851083c.md)|Ocorre quando uma janela de pasta de trabalho é ativada.|
|[WindowDeactivate](d84f0819-00df-585f-ea31-e4ab5a72950e.md)|Ocorre quando uma janela de pasta de trabalho é desativada.|
|[WindowResize](6e473482-fe16-03a2-7a27-b0cd9535c3e6.md)|Ocorre quando qualquer janela de pasta de trabalho é redimensionada.|

## Métodos



|**Nome**|**Descrição**|
|:-----|:-----|
|[AcceptAllChanges](8d8572a9-1231-c8ef-0707-72b8b5109066.md)|Aceita todas as alterações na pasta de trabalho compartilhada especificada.|
|[Activate](628e06b3-ca3f-28cb-e0fd-e696842f69f5.md)|Ativa a primeira janela associada à pasta de trabalho.|
|[AddToFavorites](14e1cd5a-41be-fc9a-61fa-df87698110e8.md)|Adiciona na pasta Favoritos um atalho para a pasta de trabalho ou para um hiperlink.|
|[ApplyTheme](11580293-22da-9154-20a0-6435b8870ac9.md)|Aplica o tema especificado à pasta de trabalho atual.|
|[BreakLink](1e9d70c1-908e-92eb-26b8-d6ac753cc9c2.md)|Converte fórmulas vinculadas a outras fontes do Microsoft Excel ou fontes OLE em valores.|
|[CanCheckIn](17f7cbdd-0ce0-8e3a-46f3-cb6dafaaa40a.md)|**True** se o Microsoft Excel pode verificar em uma pasta de trabalho especificada para um servidor. Leitura/gravação **booleano**.|
|[ChangeFileAccess](07f9cfc3-eece-efc1-6c03-38782ad7bcc2.md)|Altera as permissões de acesso para a pasta de trabalho. Isso pode exigir que uma versão atualizada seja carregada do disco.|
|[ChangeLink](9b2c0b82-73ff-3bdb-63df-82c0708cb703.md)|Altera um vínculo de um documento para outro.|
|[CheckIn](f9750086-aaa6-3b04-6b51-ebcadf6b1911.md)|Retorna uma pasta de trabalho de um computador local para um servidor e define a pasta de trabalho local como somente leitura para que ela não possa ser editada localmente. Chamar esse método também fechará a pasta de trabalho.|
|[CheckInWithVersion](3b37cea5-8795-bcbb-9c4b-d30b2b9a095e.md)|Salva uma pasta de trabalho de um computador local em um servidor e define a pasta de trabalho local como somente leitura para impedir que ela seja editada localmente.|
|[Close](c0376cab-a2db-c606-67bf-0a4921b81e03.md)|Fecha o objeto.|
|[DeleteNumberFormat](d56c2e4c-5de2-fecf-6a1f-a9fdc79943cb.md)|Exclui um formato de número personalizado da pasta de trabalho.|
|[EnableConnections](521ebb4c-56c6-3e21-39af-4a46934790e1.md)|O método  **EnableConnections** permite que os desenvolvedores habilitem programaticamente as conexões de dados na pasta de trabalho para o usuário.|
|[EndReview](cd4a445b-4731-43ba-e46a-f80f19ea5a17.md)|Finaliza a revisão de um arquivo que foi enviado para revisão usando o método  **[SendForReview](3834f5b3-6d24-1bb9-27b5-052aa2e725e3.md)**.|
|[ExclusiveAccess](9b92ec4f-e256-7e01-6cd7-759a0d022813.md)|Atribui ao usuário atual acesso exclusivo à pasta de trabalho que está aberta como uma lista compartilhada.|
|[ExportAsFixedFormat](4d72247c-bab9-3475-4792-8899c959393c.md)|O método  **ExportAsFixedFormat** é usado para publicar uma pasta de trabalho em formato PDF ou XPS.|
|[FollowHyperlink](d070ecc9-fbb6-c146-f250-5c99b09063ec.md)|Exibe um documento armazenado no cache, caso este já tenha sido baixado. Caso contrário, esse método encontra o hiperlink, baixa o documento de destino e exibe o documento no aplicativo apropriado.|
|[ForwardMailer](956b1746-26f2-5968-0ef7-fa3da2be974c.md)|Você solicitou ajuda sobre uma palavra-chave do Visual Basic usada somente no Macintosh. Para obter informações sobre essa palavra-chave, consulte a Ajuda da referência de linguagem incluída no Microsoft Office Macintosh Edition.|
|[GetWorkflowTasks](8a5ff9e0-b23a-930c-bb65-a1daa10cd946.md)|Retorna a coleção de objetos  **[WorkflowTask](9d17947e-f12a-2f97-7888-8d5ec9f85011.md)** para a pasta de trabalho especificada.|
|[GetWorkflowTemplates](adff72bb-39ab-69ed-8a9b-defe75a5fede.md)|Retorna a coleção de objetos  **[WorkflowTemplate](965d0474-dd51-9b0e-b34c-a11f921ff410.md)** para a pasta de trabalho especificada.|
|[HighlightChangesOptions](ac69ee3e-c5ea-5ac0-418a-0b94d56a8777.md)|Controla o modo como as alterações são mostradas na pasta de trabalho compartilhada.|
|[LinkInfo](016295a3-72c1-95b3-c259-8f286b12b73c.md)|Retorna o status de atualização e a data do vínculo.|
|[LinkSources](6466bea0-5af8-7af0-e9d7-7595133073ae.md)|Retorna uma matriz de vínculos da pasta de trabalho. Os nomes na matriz são os nomes dos documentos vinculados, edições ou servidores OLE ou DDE. Retorna  **vazia** se não houver nenhum link.|
|[LockServerFile](be0ac600-320e-0959-bc26-5f3f4a910f5e.md)|Bloqueia a pasta de trabalho no servidor para evitar modificação.|
|[MergeWorkbook](393790c6-3c19-7149-a999-b8712e7a6855.md)|Mescla alterações de uma pasta de trabalho em uma pasta de trabalho aberta.|
|[NewWindow](ba568cee-c1cb-6e6a-8935-2cc8ce3a8400.md)|Cria uma nova janela ou uma cópia da janela especificada.|
|[OpenLinks](cae33bab-892e-0861-e4ec-8a334097e0d1.md)|Abre os documentos de suporte de um ou mais vínculos.|
|[PivotCaches](0a2e7f10-c123-5c98-fb71-56868b9f8bde.md)|Retorna uma coleção  **[PivotCaches](cfd979b9-d52f-f34b-4b66-4fb17efcdc92.md)** que representa todas as tabelas dinâmicas armazena em cache na pasta de trabalho especificada. Somente leitura.|
|[Post](62ecf3bc-c551-8f06-64cc-a6c141bdf172.md)|Remete a pasta de trabalho especificada para uma pasta pública. Este método funciona apenas com um cliente do Microsoft Exchange conectado a um servidor do Microsoft Exchange.|
|[PrintOut](3a4e7037-fcde-5a57-4a80-45f2a0994370.md)|Imprime o objeto.|
|[PrintPreview](044afc4c-74d6-3ea6-1811-2c7d9cdc5b1a.md)|Mostra uma visualização de como o objeto aparecerá quando for impresso.|
|[Protect](0e270b93-7b0b-cc68-c7c0-4002024f4292.md)|Protege uma pasta de trabalho de forma que ela não possa ser modificada.|
|[ProtectSharing](26660bc6-136a-ffc8-987e-c96db9c08231.md)|Salva a pasta de trabalho a protege para compartilhamento.|
|[PurgeChangeHistoryNow](7ea42af1-051b-400d-cb87-0736c49d74fb.md)|Remove entradas do registro de alterações da pasta de trabalho especificada.|
|[RefreshAll](c1a956dc-263c-5c24-3b51-fc4af22dcd33.md)|Atualiza todos os intervalos de dados externos e os relatórios da tabela dinâmica na pasta de trabalho especificada.|
|[RejectAllChanges](a53670da-370c-9ac8-75b8-008625495c2b.md)|Rejeita todas as alterações em uma pasta de trabalho compartilhada especificada.|
|[ReloadAs](ce6a9d1a-7945-3dca-ff2d-a42289c2ccf9.md)|Recarrega uma pasta de trabalho com base em um documento HTML, usando a codificação de documento especificada.|
|[RemoveDocumentInformation](e668d976-108b-c627-6118-dd3384c1315c.md)|Remove todas as informações do tipo especificado da pasta de trabalho.|
|[RemoveUser](f0a978a0-7bcf-3af4-a01a-831c6c854989.md)|Desconecta o usuário especificado da pasta de trabalho compartilhada.|
|[Reply](557bb3a4-c817-e942-10cf-ba252b0db498.md)|Você solicitou ajuda sobre uma palavra-chave do Visual Basic usada somente no Macintosh. Para obter informações sobre essa palavra-chave, consulte a Ajuda da referência de linguagem incluída no Microsoft Office Macintosh Edition.|
|[ReplyAll](c378da35-1778-44db-5c58-8d6992ca0c93.md)|Você solicitou ajuda sobre uma palavra-chave do Visual Basic usada somente no Macintosh. Para obter informações sobre essa palavra-chave, consulte a Ajuda da referência de linguagem incluída no Microsoft Office Macintosh Edition.|
|[ReplyWithChanges](60424d69-0062-aa5e-ea8f-4fb07086167a.md)|Envia uma mensagem de email para o autor de uma pasta de trabalho que foi enviada para revisão, notificando que um revisor concluiu a revisão da pasta de trabalho.|
|[ResetColors](1b56a4e9-0645-fa1e-55cc-09069c6a0ff1.md)|Redefine a paleta de cores com as cores padrão.|
|[RunAutoMacros](85dfdadf-75e6-437d-fb7a-e17681a69b35.md)|Executa a macro Auto_Open, Auto_Close, Auto_Activate ou Auto_Deactivate anexada à pasta de trabalho. Esse método é incluído para compatibilidade com versões anteriores. Para código novo do Visual Basic, você deve usar os eventos Open, Close, Activate e Deactivate em vez dessas macros.|
|[Save](466d891d-fb4c-fb0a-474b-dedb3c4ea7a7.md)|Salva alterações na pasta de trabalho especificada.|
|[SaveAs](fbc3ce55-27a3-aa07-3fdb-77b0d611e394.md)|Salva alterações feitas na pasta de trabalho em um arquivo diferente.|
|[SaveAsXMLData](7c4c1be3-d3a5-6e90-7750-9f371f008541.md)|Exporta os dados, que foram mapeados para o mapa de esquema XML especificado, para um arquivo de dados XML.|
|[SaveCopyAs](84f58488-6a2b-7fef-1472-e1b9771a60b0.md)|Salva uma cópia da pasta de trabalho em um arquivo mas não modifica a pasta de trabalho aberta na memória.|
|[SendFaxOverInternet](e7d91ac4-90d2-7555-af96-dc28736da769.md)|Envia uma planilha como um fax para os destinatários especificados.|
|[SendForReview](3834f5b3-6d24-1bb9-27b5-052aa2e725e3.md)|Envia uma pasta de trabalho em uma mensagem de email para revisão aos destinatários especificados.|
|[SendMail](581d197c-0748-2225-2986-64aa368aab39.md)|Envia a pasta de trabalho usando o sistema de email instalado.|
|[SendMailer](e44955e1-e250-7279-19e5-e13db80ceddc.md)|Você solicitou ajuda sobre uma palavra-chave do Visual Basic usada somente no Macintosh. Para obter informações sobre essa palavra-chave, consulte a Ajuda da referência de linguagem incluída no Microsoft Office Macintosh Edition.|
|[SetLinkOnData](b500a579-6e4c-5712-05cf-27c6393b3bcd.md)|Define o nome de um procedimento que é executado sempre que um vínculo DDE é atualizado.|
|[SetPasswordEncryptionOptions](3b6c9bfe-4cfb-1dde-fd57-07dd474df7db.md)|Define as opções para criptografia de pastas de trabalho usando senhas.|
|[ToggleFormsDesign](3a6352e3-26b9-713e-ed93-a5890b37bc0a.md)|O método  **ToggleFormsDesign** é usado para alternar o Excel no modo de Design ao usar controles de formulários.|
|[Unprotect](39387902-a8a4-7bf2-44d7-c5bde6725778.md)|Remove a proteção de uma planilha ou pasta de trabalho. Este método não tem efeito algum se a planilha ou pasta de trabalho não estiver protegida.|
|[UnprotectSharing](edce1744-0906-4b4e-8b98-5d1125047bff.md)|Desativa a proteção para o compartilhamento e salva a pasta de trabalho.|
|[UpdateFromFile](f5148b60-9b25-8a12-5cf3-40103dcff2a3.md)|Atualiza uma pasta de trabalho somente leitura a partir de sua versão salva em disco se esta for mais recente que a cópia da pasta de trabalho carregada na memória. Se a cópia em disco não tiver sido alterada desde que a pasta de trabalho foi carregada, a cópia em memória da pasta de trabalho não será recarregada.|
|[UpdateLink](2aef72cc-a820-3e91-1f46-50c739faf2bb.md)|Atualiza um vínculo (ou vínculos) DDE, OLE ou do Microsoft Excel.|
|[WebPagePreview](2c88f15e-5cd3-82da-f779-55b63959a2b0.md)|Exibe uma visualização da aparência da pasta de trabalho especificada como se tivesse sido salva como uma página da Web.|
|[XmlImport](97964c82-1fbe-7060-0a90-23c190e0b398.md)|Importa um arquivo de dados XML para a pasta de trabalho atual.|
|[XmlImportXml](b0edbe49-f578-ead0-8371-0196f5c515d4.md)|Importa um fluxo de dados XML que tenha sido previamente carregado na memória. O Excel usa o primeiro mapa de qualificação encontrado ou, se o intervalo de destino for especificado, o Excel listará automaticamente os dados.|
|[CreateForecastSheet](bec7b60b-7840-af15-6d5f-f5c184ea7aee.md)|Se você tiver dados históricos baseadas em tempo, você pode usar **CreateForecastSheet**para criar uma previsão. Quando você cria uma previsão, uma nova planilha é criada que contém uma tabela dos valores históricos e previstos e um gráfico mostrando a isso. Uma previsão pode ajudá-lo a prever coisas como vendas no futuro, requisitos de inventário ou tendências de consumidor.|

## Propriedades



|**Nome**|**Descrição**|
|:-----|:-----|
|[AccuracyVersion](bc81782c-662c-87ec-8381-d06e77674d0c.md)|Especifica se algumas funções de planilha usam os mais recentes algoritmos de precisão para calcular seus resultados. Leitura/gravação.|
|[ActiveChart](81e18252-b1fe-2487-535e-6e24c80bef24.md)|Retorna um objeto  **[Chart](179c32ce-49bd-6f36-ea12-89fb5443f3ea.md)** representando o gráfico ativo (um gráfico incorporado ou uma planilha de gráfico). Um gráfico incorporado é considerado ativo quando ele está selecionado ou ativado. Quando nenhum gráfico está ativo, essa propriedade retornará **Nothing**.|
|[ActiveSheet](fb5578c3-64a7-edb7-4004-e608739d4c1e.md)|Retorna um objeto que representa a planilha ativa (a planilha na parte superior) na pasta de trabalho ativa ou na janela especificada ou pasta de trabalho. Retorna  **Nothing** se não planilha ativa.|
|[ActiveSlicer](d3858353-0be1-338c-e43f-1e5ffb7f37ba.md)|Retorna um objeto que representa a segmentação de dados ativa na pasta de trabalho ativa ou na pasta de trabalho especificada. Retorna  **Nothing** se nenhum segmentação de dados estiver ativa. Somente leitura.|
|[Application](91b30f9d-48e5-e033-8daf-416d1c0e547d.md)|Quando usado sem um qualificador de objeto, essa propriedade retorna um objeto  **[Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** que representa o aplicativo Microsoft Excel. Quando usado com um qualificador de objeto, essa propriedade retorna um objeto **Application** que representa o criador do objeto especificado (você pode usar essa propriedade com um objeto de automação OLE para retornar o aplicativo desse objeto). Somente leitura.|
|[AutoUpdateFrequency](dfded5c8-94d6-8a0f-24c1-248bd502850b.md)|Retorna ou define o número de minutos entre as atualizações automáticas para a pasta de trabalho compartilhada. Leitura/gravação  **longa**.|
|[AutoUpdateSaveChanges](06f9951d-a17a-bf88-4f6e-65835eb112f8.md)|**True** se as alterações atuais da pasta de trabalho compartilhada são remetidas para outros usuários sempre que a pasta de trabalho é atualizada automaticamente. **False** se as alterações não forem remetidas (nesta pasta de trabalho ainda está sincronizada com as alterações feitas por outros usuários). O valor padrão é **True**. Leitura/gravação **booleano**.|
|[BuiltinDocumentProperties](3efffd7d-0681-ecbc-000a-b71eceb3f92a.md)|Retorna uma coleção  **[DocumentProperties](90d42786-7d9a-b604-dbdf-88db41cbe69b.md)** que representa todas as propriedades de documento internas para a pasta de trabalho especificada. Somente leitura.|
|[CalculationVersion](09633164-998f-9fa7-f257-da109c369cd7.md)|Retorna as informações sobre a versão do Excel que a pasta de trabalho foi totalmente recalculada pela última vez. Somente leitura  **Long**.|
|[CaseSensitive](6053b576-9ede-f9d8-e2bf-c012653b60a2.md)|**True** se a pasta de trabalho faz distinção entre maiúsculas e minúsculas durante a comparação de conteúdo. Somente leitura **booleano**|
|[ChangeHistoryDuration](5ebc3cc5-dffa-60cf-08cb-b2f84424c4b4.md)|Retorna ou define o número de dias mostrado no histórico de alterações da pasta de trabalho compartilhada. Leitura/gravação  **longa**.|
|[ChartDataPointTrack](0aa2b1c1-0bba-f514-6158-00cdb4a5747e.md)|**True** fará com que todos os gráficos no documento atual para rastrear o valor real de ponto de dados à qual está associada. **False** reverterá para o índice do ponto de dados de acompanhamento. **Boolean** Leitura/gravação|
|[Charts](582d9a78-d86f-ab69-0c22-85f8a59412d9.md)|Retorna uma coleção  **[Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** que representa todas as planilhas de gráfico da pasta de trabalho especificada.|
|[CheckCompatibility](9379c010-6756-b7ea-b4ad-5c8a4b900124.md)|Controla se o verificador de compatibilidade é executado automaticamente ou não quando a pasta de trabalho é salva.  **Boolean** de leitura/gravação.|
|[CodeName](236e97b8-2bb9-c3a9-b4da-b1c327acde95.md)|Retorna o nome de código para o objeto. Somente leitura  **cadeia de caracteres**.|
|[Colors](60fc038b-980b-c1bc-6d1c-69d9d31a11ba.md)|Retorna ou define as cores na paleta para a pasta de trabalho. A paleta tem 56 entradas, cada um representado por um valor RGB. Leitura/gravação  **Variant**.|
|[CommandBars](8d93b8cd-c4e3-b216-eda0-da4c6e573c40.md)|Retorna um objeto  **[CommandBars](0e312e21-14ee-5055-d604-b66e61c53b47.md)** que representa as barras de comandos do Microsoft Excel. Somente leitura.|
|[ConflictResolution](5142c848-0731-14d9-5913-bbaa67bf308f.md)|Retorna ou define a maneira como os conflitos deverão ser resolvidos sempre que uma pasta de trabalho compartilhada for atualizada.  **[XlSaveConflictResolution](1cdccb5a-c356-4572-9429-49850338b31b.md)** Boolean de leitura/gravação.|
|[Connections](9c4f4ba7-dd4b-0bc2-65b7-16455014097f.md)|A propriedade  **Connections** estabelece uma conexão entre a pasta de trabalho e um ODBC ou uma fonte de dados OLEDB e atualiza os dados sem avisar o usuário. Somente leitura.|
|[ConnectionsDisabled](afd53cc5-12d8-4b22-3186-1359c14f662e.md)|Desabilita os links ou conexões externas na pasta de trabalho. Somente leitura|
|[Container](7ad370bc-9901-3b8b-12e6-1ee57f0300e0.md)|Retorna o objeto que representa o aplicativo de contêiner para o objeto OLE especificado. Somente leitura  **objeto**.|
|[ContentTypeProperties](a2919232-3fa2-7f37-00c2-48eb3edb10fd.md)|Retorna uma coleção  **[MetaProperties](957a6e06-3348-b180-3655-06ffbfb69e12.md)** que descreve os metadados armazenados na pasta de trabalho. Somente leitura.|
|[CreateBackup](33f05bf8-00ef-81f4-c083-30326f019cd4.md)|**True** se um arquivo de backup for criado quando este arquivo for salvo. Somente leitura **booleano**.|
|[Creator](e03bdff2-7a93-f882-31a1-1ba8dd3c1764.md)|Retorna um inteiro de 32 bits que indica o aplicativo no qual esse objeto foi criado. Somente leitura  **Long**.|
|[CustomDocumentProperties](8470adbb-5b10-96ba-71f7-c667c33b6707.md)|Retorna ou define uma coleção  **[DocumentProperties](90d42786-7d9a-b604-dbdf-88db41cbe69b.md)** que representa todas as propriedades de documento personalizadas da pasta de trabalho especificada.|
|[CustomViews](286f6d5a-fb91-a339-8e74-9014ab7f4835.md)|Retorna uma coleção  **[CustomViews](f970bdf7-371b-ba41-89a3-bef2c6907f1a.md)** representando todas as exibições personalizadas da pasta de trabalho.|
|[CustomXMLParts](bd31f001-0e5d-691b-e69e-4cb91a6dbb0e.md)|Retorna uma coleção  **[CustomXMLParts](98c1c58e-a08d-6304-8626-1e6705917da3.md)** que representa o XML personalizado no repositório de dados XML. Somente leitura.|
|[Date1904](0556311c-4e45-aea3-e922-24a5830b19d4.md)|**True** se a pasta de trabalho usa o sistema de datas 1904. Leitura/gravação **booleano**.|
|[DefaultPivotTableStyle](8e2ca78a-4eb1-1b1e-c947-8a724f6d690a.md)|Especifica o estilo de tabela da coleção  **TableStyles** que é usado como o estilo padrão para tabelas dinâmicas. Leitura/gravação.|
|[DefaultSlicerStyle](0f193fb8-b766-9093-9db8-8b028da108b4.md)|Especifica o estilo do objeto  **[TableStyles](952da370-51cb-b1e0-a413-15cb558099b5.md)** que é usado como o estilo padrão das segmentações de dados. Leitura/gravação.|
|[DefaultTableStyle](2dc86b2c-0047-53b5-3cc4-af15c36b78cf.md)|Especifica o estilo de tabela da coleção  **TableStyles** que é usado como o padrão TableStyle. Leitura/gravação **Variant**.|
|[DefaultTimelineStyle](78261166-759a-8b18-c194-1f9124ca7e4e.md)|O nome do estilo padrão segmentação de dados da pasta de trabalho. **Variante**Leitura/gravação|
|[DisplayDrawingObjects](78eec8af-416d-88e6-d1f4-0b97a008f752.md)|Retorna ou define como as formas são exibidas. Leitura/gravação  **longa**.|
|[DisplayInkComments](bce6b184-7640-f51c-1feb-1973de6ff739.md)|Um valor  **Boolean** que determina se os comentários sobre tinta é exibidos na pasta de trabalho. Leitura/gravação **booleano**.|
|[DocumentInspectors](26d2575f-6e61-4509-6a67-45ae576bc9fe.md)|Retorna uma coleção  **[DocumentInspectors](8366d7cd-e016-bb99-d27f-749ca10352f1.md)** que representa os módulos do Inspetor de documentos para a pasta de trabalho especificada. Somente leitura.|
|[DocumentLibraryVersions](b6338994-b5d9-ef9b-83b5-bdd47d0fd407.md)|Retorna uma coleção  **[DocumentLibraryVersions](075c0315-fade-6d45-9ab9-6c798f6f09ac.md)** que representa a coleção de versões de uma pasta de trabalho compartilhada com versão habilitada e que está armazenada em uma biblioteca de documentos em um servidor.|
|[DoNotPromptForConvert](d2af6528-4d9f-6e94-4fa6-2322098b4b17.md)|Retorna ou define se o usuário deve ser solicitado a converter a pasta de trabalho, se a pasta de trabalho contém recursos que não são suportados pelo versões do Excel anteriores ao Excel 2007. Leitura/gravação  **booleano**.|
|[EnableAutoRecover](04a82e4d-0231-adf1-1289-35514372c995.md)|Salva arquivos alterados, de todos os formatos, em um intervalo cronometrado. Leitura/gravação  **booleano**.|
|[EncryptionProvider](13047af7-2e6e-6b64-05f1-8b4bd7a734ad.md)|Retorna uma  **cadeia de caracteres** especificando o nome do provedor de criptografia de algoritmo que o Microsoft Office Excel 2007 usa ao criptografar documentos. Leitura/gravação.|
|[EnvelopeVisible](d511a75a-ddd1-64f5-a09b-720657f64c09.md)|**True** se o cabeçalho de composição de email e a barra de ferramentas de envelope estiverem visíveis. Leitura/gravação **booleano**.|
|[Excel4IntlMacroSheets](70a8c8d0-1169-7c3d-904e-5a32a4693f45.md)|Retorna uma coleção  **[Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** que representa todas as folhas de macro internacionais do Microsoft Excel 4.0 na pasta de trabalho especificada. Somente leitura.|
|[Excel4MacroSheets](29161ab8-da75-c7b5-561d-f4423b8ab1ef.md)|Retorna uma coleção  **[Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** que representa todas as folhas de macro do Microsoft Excel 4.0 na pasta de trabalho especificada. Somente leitura.|
|[Excel8CompatibilityMode](8471493b-8733-cddf-75fa-42d3d1719300.md)|A propriedade  **Excel8CompatibilityMode** fornece aos desenvolvedores uma maneira de verificar se a pasta de trabalho está no modo de compatibilidade. Somente leitura **booleano**.|
|[FileFormat](ef722c3c-90ea-9810-b853-a3fff19d5c60.md)|Retorna o formato de arquivo e/ou o tipo da pasta de trabalho. Somente leitura  **[XlFileFormat](4c0ebc4c-915c-c199-ee39-f4d14ba7b64e.md)**.|
|[Final](55d3a155-ca0c-1f7c-8612-80aac91a8eb3.md)|Retorna ou define um  **Boolean** que indica se uma pasta de trabalho é final. Leitura/gravação **booleano**.|
|[ForceFullCalculation](76f46d18-79e3-9828-d126-e221ae1a8157.md)|Retorna ou define a pasta de trabalho especificada para o modo de cálculo forçado. Leitura/gravação.|
|[FullName](83f45d15-b009-f304-ca53-4daa80c06562.md)|Retorna o nome do objeto, incluindo seu caminho no disco, como uma cadeia de caracteres. Somente leitura  **cadeia de caracteres**.|
|[FullNameURLEncoded](589d98f7-e6fa-bc28-2c8f-7cb72009737a.md)|Retorna uma  **cadeia de caracteres** indicando o nome do objeto, incluindo seu caminho no disco, como uma cadeia de caracteres. Somente leitura.|
|[HasPassword](e3cfdc90-1e82-5556-0064-e8269ba92539.md)|**True** se a pasta de trabalho tem uma senha de proteção. Somente leitura **booleano**.|
|[HasVBProject](b4293266-40d9-a8a4-80ff-8b19ec7ed823.md)|Retorna um  **Boolean** que indica se uma pasta de trabalho tem um anexado Microsoft projeto Visual Basic for Applications. Somente leitura **booleano**.|
|[HighlightChangesOnScreen](146f9a16-d32b-cc8f-fece-03864f0e13a2.md)|**True** se as alterações na pasta de trabalho compartilhada são destacadas na tela. Leitura/gravação **booleano**.|
|[IconSets](c837d2a8-d21d-7432-a409-f49426368556.md)|Essa propriedade é usada para filtrar os dados em uma pasta de trabalho com base em um ícone de célula da coleção  **IconSet**. Somente leitura.|
|[InactiveListBorderVisible](a6259862-9a29-f3a5-498f-633f51ec10e6.md)|Um valor  **Boolean** que especifica se as bordas da lista são visíveis quando uma lista não estiver ativa. Retorna **True** se a borda estiver visível. Leitura/gravação **booleano**.|
|[IsAddin](b8c8b9f4-4be5-0260-957e-c6450f31a0c0.md)|**True** se a pasta de trabalho está sendo executado como um add-in. leitura/gravação **booleano**.|
|[IsInplace](f492c09f-79d1-cde0-6cf1-db9644e41589.md)|**True** se a pasta de trabalho especificada estiver sendo editada no lugar. **False** se a pasta de trabalho tiver sido aberta no Microsoft Excel para edição. Somente leitura **booleano**.|
|[KeepChangeHistory](3dbc322e-2b93-ae3d-cb9e-64c657fc0f80.md)|**True** se o controle de alterações estiver habilitada para a pasta de trabalho compartilhada. Leitura/gravação **booleano**.|
|[ListChangesOnNewSheet](77adf429-baa5-f2be-6139-c2b07dda5174.md)|**True** se as alterações na pasta de trabalho compartilhada são mostradas em uma planilha separada. Leitura/gravação **booleano**.|
|[Mailer](b020d3f6-7120-d03c-bc42-c297bcfbebf6.md)|Você solicitou ajuda sobre uma palavra-chave do Visual Basic usada somente no Macintosh. Para obter informações sobre essa palavra-chave, consulte a Ajuda da referência de linguagem incluída no Microsoft Office Macintosh Edition.|
|[Model](43ccdaa8-4a12-e745-88db-9db8a328ee5e.md)|Retorna o objeto de  **modelo**, que é um modelo de dados da pasta de trabalho de nível superior. Somente leitura|
|[MultiUserEditing](dc721463-ec34-8c52-6701-51c406beed23.md)|**True** se a pasta de trabalho está aberta como uma lista compartilhada. Somente leitura **booleano**.|
|[Name](55526a99-da9c-7f14-55f7-dfe9bd8ff489.md)|Retorna um valor  **String** que representa o nome do objeto.|
|[Names](26be56ec-ea12-1600-602a-eb338d4a5a8b.md)|Retorna uma coleção de  **[nomes](ffecf89d-7bae-c470-8e37-608857a9de2a.md)** que representa todos os nomes na pasta de trabalho especificada (incluindo todos os nomes de planilha específico). Objeto de **nomes** de somente leitura.|
|[Parent](4c039b5b-050f-8f4d-bc90-7982e92fb17c.md)|Retorna o objeto pai do objeto especificado. Somente leitura.|
|[Password](5eaaf8cd-4344-946e-ecfa-c0f48946d2f2.md)|Retorna ou define a senha que deve ser fornecida para abrir a pasta de trabalho especificada.  **Cadeia de caracteres** de leitura/gravação.|
|[PasswordEncryptionAlgorithm](2745a8da-2a61-b949-115a-7f1112a0289e.md)|Retorna uma  **cadeia de caracteres** indicando o algoritmo que o Microsoft Excel usa para criptografar senhas para a pasta de trabalho especificada. Somente leitura.|
|[PasswordEncryptionFileProperties](536ad729-424e-5f81-85e9-8a6ed71fb11a.md)|**True** se o Microsoft Excel criptografa as propriedades de arquivo para a pasta de trabalho especificada protegida por senha. Somente leitura **booleano**.|
|[PasswordEncryptionKeyLength](2662f2f5-1ad0-4a75-82c0-3268f147948a.md)|Retorna um  **Long** indicando o comprimento da chave do algoritmo do Microsoft Excel usa ao criptografar senhas para a pasta de trabalho especificada. Somente leitura.|
|[PasswordEncryptionProvider](d5bcbbf2-8de9-6725-9cac-679d6c023b34.md)|Retorna uma  **cadeia de caracteres** especificando o nome do provedor de criptografia de algoritmo que o Microsoft Excel usa ao criptografar senhas para a pasta de trabalho especificada. Somente leitura.|
|[Path](f4cbf76a-2ed3-63b7-3262-45403d6f086e.md)|Retorna uma  **cadeia de caracteres** que representa o caminho completo para o arquivo/pasta de trabalho que esta pasta de trabalho do objeto representa.|
|[Permission](ef04f56e-a04d-c3d9-fdda-611be7bf9d39.md)|Retorna um objeto  **Permission** que representa as configurações de permissão na pasta de trabalho especificada.|
|[PersonalViewListSettings](998320bf-d703-e42f-8b43-5a7b909a846d.md)|**True** se as configurações de filtro e classificação para listas estão incluídas na exibição de pessoal do usuário da pasta de trabalho compartilhada. Leitura/gravação **booleano**.|
|[PersonalViewPrintSettings](6e4a0a9c-4eb0-d8e1-e9ce-8e9e618996b4.md)|**True** se as configurações de impressão são incluídas no modo de exibição pessoal do usuário da pasta de trabalho compartilhada. Leitura / gravação **booleano**.|
|[PivotTables](b11795e0-22c8-f089-c59a-5e3d7a09d5de.md)|Retorna um object que representa uma coleção de todas as tabelas dinâmicas relatórios em uma planilha. Somente leitura.|
|[PrecisionAsDisplayed](4f0c8201-5b8d-5cb5-337c-944d2c7dd8d1.md)|**True** se cálculos nesta pasta de trabalho serão feitos usando apenas a precisão dos números como exibidos. Leitura/gravação **booleano**.|
|[ProtectStructure](bf721b60-0ad1-f71c-7ef4-74d2196d320e.md)|**True** se a ordem das planilhas na pasta de trabalho está protegida. Somente leitura **booleano**.|
|[ProtectWindows](0f285fbe-2545-5c7d-9e3d-f08d57e78092.md)|**True** se as janelas da pasta de trabalho estiverem protegidas. Somente leitura **booleano**.|
|[PublishObjects](b6418f80-5154-6e3f-7313-222e6438c0e1.md)|Retorna a coleção  **[PublishObjects](33ad393e-5ab6-2531-5e5b-42930fc596c0.md)**. Somente leitura.|
|[ReadOnly](f3c0ec74-63af-ed76-f854-ce2382b9fcf3.md)|Retorna  **True** se o objeto tiver sido aberto como somente leitura. Somente leitura **booleano**.|
|[ReadOnlyRecommended](3cae84e4-d5f0-f01c-64d9-ec586ffdf79c.md)|**True** se a pasta de trabalho foi salvo como somente leitura recomendado. Somente leitura **booleano**.|
|[RemovePersonalInformation](f5cdc655-8ba9-6dd1-ab05-028d98c11972.md)|**True** se as informações pessoais podem ser removidas da pasta de trabalho especificada. O valor padrão é **False**. Leitura/gravação **booleano**.|
|[Research](3a7ba740-314b-664b-3be6-1e8cdeded234.md)|Retorna um objeto  **Research** que representa o serviço de pesquisa para uma pasta de trabalho. Somente leitura.|
|[RevisionNumber](7ea9fde5-eb89-a9b0-b637-980f1533d8ec.md)|Retorna o número de vezes que a pasta de trabalho foi salva enquanto aberta como uma lista compartilhada. Se a pasta de trabalho for aberta no modo exclusivo, essa propriedade retorna 0 (zero). Somente leitura  **Long**.|
|[Saved](37eb8e08-2bfa-8065-2520-a71e291ab50c.md)|**True** se não tiver sofrido alterações na pasta de trabalho especificada desde que foi salvo pela última vez. Leitura/gravação **booleano**.|
|[SaveLinkValues](ee69911f-5a4a-5c2b-c14a-cd562f3ba9f4.md)|**True** se o Microsoft Excel salvar valores de vínculo externos com a pasta de trabalho. Leitura/gravação **booleano**.|
|[ServerPolicy](188f6c47-35e3-bb69-cb8d-9d78b5b8fea5.md)|Retorna um objeto  **ServerPolicy** que representa uma diretiva especificada para uma pasta de trabalho armazenada em um servidor executandoSharePoint Server2007 ou posterior. Somente leitura.|
|[ServerViewableItems](2c10a647-2b2c-0640-9990-109b89444cd2.md)|Permite que um desenvolvedor interaja com a lista de objetos publicados na pasta de trabalho que são mostrados no servidor. Somente leitura.|
|[SharedWorkspace](864fdee9-7149-360f-099d-e1a9b57a31db.md)|Este objeto ou membro foi substituído, mas continua a fazer parte do modelo de objeto para compatibilidade com versões anteriores. Você não deve usá-lo nos novos aplicativos.|
|[Sheets](45e4e19e-55ea-9615-231d-9435ba6d5a63.md)|Retorna uma coleção  **[Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** que representa todas as planilhas na pasta de trabalho especificada. Objeto **Sheets** somente leitura.|
|[ShowConflictHistory](d8588b9e-3e4b-6224-aaa7-ce0b63ff0607.md)|**True** se a planilha de histórico de conflito estiver visível na pasta de trabalho que está aberta como uma lista compartilhada. Leitura/gravação **booleano**.|
|[ShowPivotChartActiveFields](8892b134-4882-e1ff-a265-65b36af66f1a.md)|Essa propriedade controla a visibilidade do painel de filtro gráfico dinâmico. Leitura/gravação  **booleano**.|
|[ShowPivotTableFieldList](33c74c54-27ea-d230-c640-47109bdfd4a2.md)|**True** (padrão) se a lista de campos da tabela dinâmica pode ser exibida. Leitura/gravação **booleano**.|
|[Signatures](b45f8036-c2d7-6113-e95c-ff78ee6a1f46.md)|Retorna as assinaturas digitais de uma pasta de trabalho. Somente leitura.|
|[SlicerCaches](1ebb7fd1-1742-815a-b4bb-4d25d6c9e705.md)|Retorna o objeto  **[SlicerCaches](d6097f70-cdc7-3be7-575c-cf43a0765e10.md)** associado à pasta de trabalho. Somente leitura.|
|[SmartDocument](19916b63-e93a-7f1e-532c-f4bbdb60622d.md)|Retorna um objeto  **SmartDocument** que representa as configurações para uma solução de documento inteligente. Somente leitura.|
|[Styles](c9a70be9-cab5-ea5f-2e3f-949b1acf43d9.md)|Retorna uma coleção de  **[estilos](146effdc-e007-814d-b110-f7bd944fc15f.md)** que representa todos os estilos na pasta de trabalho especificada. Somente leitura.|
|[Sync](000c9739-13ab-d6eb-c1c3-2ce721911137.md)|Este objeto ou membro foi substituído, mas continua a fazer parte do modelo de objeto para compatibilidade com versões anteriores. Você não deve usá-lo nos novos aplicativos.|
|[TableStyles](ac23db99-b2ce-0228-7808-728817736694.md)|Retorna um objeto da coleção  **TableStyles** da pasta de trabalho atual referente aos estilos usados na pasta de trabalho atual. Somente leitura.|
|[TemplateRemoveExtData](9851df1d-4e83-525a-8a43-bd84b0a94c74.md)|**True** se as referências de dados externos são removidas quando a pasta de trabalho é salva como um modelo. Leitura/gravação **booleano**.|
|[Theme](1208f610-8c6f-9a62-3378-9566a7ee6b37.md)|Retorna o tema aplicado à pasta de trabalho atual. Somente leitura.|
|[UpdateLinks](c8d374d7-0b32-eb32-fa29-ab496d6786e7.md)|Retorna ou define uma constante  **[XlUpdateLink](8ddd9876-7c24-09dd-5b89-33804adc2097.md)** indicando a configuração da pasta de trabalho para atualizar vínculos OLE incorporados. Leitura/gravação.|
|[UpdateRemoteReferences](055c1a88-c189-ddd3-c9b2-9458817cec90.md)|**True** se o Microsoft Excel atualiza referências remotas na pasta de trabalho. Leitura/gravação **booleano**.|
|[UserStatus](0df24f8a-b60b-fd8c-3436-903652487a09.md)|Retorna uma matriz bidimensional, baseado em 1 que fornece informações sobre cada usuário com a pasta de trabalho aberta como uma lista compartilhada. Somente leitura  **Variant**.|
|[UseWholeCellCriteria](b65093aa-37ca-2aa1-4f18-c90bc7536f74.md)|**True** se a pasta de trabalho usa os padrões de pesquisa que correspondem a todo o conteúdo de uma célula. Somente leitura **booleano**.|
|[UseWildcards](92e7463c-6dbe-c409-461a-ca730402ad62.md)|**True** se a pasta de trabalho permite que os caracteres curinga para comparações de sequência de caracteres e pesquisa. Somente leitura **booleano**|
|[VBASigned](6e93161c-2fa4-1064-9b5d-a8eb96ad2bea.md)|**True** se o Visual Basic for Applications do projeto para a pasta de trabalho especificada foi assinado digitalmente. Somente leitura **booleano**.|
|[VBProject](1bef5b7e-e169-fa4b-9810-6cd87ecd0a8d.md)|Retorna um objeto  **VBProject** que representa o projeto do Visual Basic na pasta de trabalho especificada. Somente leitura.|
|[WebOptions](801742a2-f5d8-5311-ea24-fd428532ba80.md)|Retorna a coleção  **[WebOptions](d573637f-1891-4602-c961-091795e47356.md)**, que contém atributos no nível de pasta de trabalho usados pelo Microsoft Excel quando você salva um documento como uma página da Web ou abre uma página da Web. Somente leitura.|
|[Windows](2352d6c9-720e-b58d-6e7c-049bf21a090d.md)|Retorna uma coleção  **[Windows](d5d0e3c9-9132-469c-d033-d29397dacd77.md)** que representa todas as janelas na pasta de trabalho especificada. Objeto **Windows** somente leitura.|
|[Worksheets](8b7d660d-ca49-0bd0-dc57-64defa47bd5e.md)|Retorna uma coleção  **[Sheets](048fd93c-bc27-4b58-358f-56fcee1710f8.md)** que representa todas as planilhas da pasta de trabalho especificada. Objeto **Sheets** somente leitura.|
|[WritePassword](ac89063e-6ef5-f7c5-abb0-4e6ef1c5fd05.md)|Retorna ou define uma  **cadeia de caracteres** para a senha de gravação de uma pasta de trabalho. Leitura/gravação.|
|[WriteReserved](96cc86d1-0e77-b6f3-3045-f6346de0f969.md)|**True** se a pasta de trabalho é reservado para gravação. Somente leitura **booleano**.|
|[WriteReservedBy](f053c197-3af3-9ab7-bee1-f72ee311a5b8.md)|Retorna o nome do usuário que atualmente tem permissão de gravação da pasta de trabalho. Somente leitura  **cadeia de caracteres**.|
|[XmlMaps](c7893167-bfa1-e1df-58f3-782b80322fad.md)|Retorna uma coleção  **XmlMaps** que representa os mapas de esquema que foram adicionados à pasta de trabalho especificada. Somente leitura.|
|[XmlNamespaces](b93aba02-f831-6321-1c0d-2a30d417e57f.md)|Retorna uma coleção  **[XmlNamespaces](430f6773-2be5-8312-cd67-afb703ab0782.md)** que representa os namespaces XML contidos na pasta de trabalho especificada. Somente leitura.|
|[Queries](29ee16cb-b6f2-2358-7e1a-3b1e7f9cf654.md)||
