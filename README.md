# SPConvertWord2Pdf
SPConvertWord2Pdf é um projeto cujo o cenário é o seguinte: Será subido um documento nas seguintes extensões (*.doc) ou (*.docx) para biblioteca [Documento Compartilhados] do Sharepoint, esta biblioteca terá um campo chamado [Convert] do tipo booleano e serve para saber se será preciso criar um arquivo pdf do original.
Esta biblioteca tem também o campo [StatusConversao] que podem ser {NotStarted, InProgress, Canceled, Failed, Succeeded}, campo [Informacao] que serve para escrever as informações caso tenha falha ou cancelamento da conversão. O campo [StartTime] guarda o início da conversão e o campo [CompleteTime] guarda o término da conversão e ambos só são preenchidos caso a conversão tenha sido um sucesso.

O novo arquivo (*.pdf) será posto na biblioteca [DocPdf].

Para tal realização será utilizado a library [Microsoft.Office.Word.Server] que se encontra em ["C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.Office.Word.Server.dll"] que é responsável pelo serviço de Automatização do Word no Sharepoint 2013.

E nele fazemos uso da classe SyncConverter, fazendo solicitação com base imadiata, ou seja, sem espera;
executando conversão imediatamente, não necessáriamente do TimerJob; opera no arquivo no momento da solicitação; podendo executar conversões de arquivos e atualizações de arquivos; também notifica ou atualiza itens no Sharepoint após a conclusão.

Para mais informações sobre a classe SyncConverter, consulte https://msdn.microsoft.com/en-us/library/microsoft.office.conversionservices.conversions.syncconverter.aspx