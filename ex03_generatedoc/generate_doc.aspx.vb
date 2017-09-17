Imports Aspose.Words
Imports Aspose.Words.Tables

Public Class generate_doc
    Inherits System.Web.UI.Page

    Protected builder As DocumentBuilder = Nothing

    Private Sub btnGen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGen.Click
        builder = New DocumentBuilder()

        SetDocPage()

        WriteDocHeaderAndFooter()
        WriteDocParagraphs()
        WriteDocTextSeparately()
        WriteDocList()
        WriteDocTable()

        SaveDocFile()
    End Sub

    ''' <summary>
    ''' 頁面相關設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetDocPage()
        ' How-to: Change Page Setup for Whole Document
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/howto-change-page-setup-for-whole-document.html
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/aspose.words.section.pagesetup.html
        Dim pageSetup As PageSetup = builder.PageSetup

        With pageSetup
            .PaperSize = PaperSize.A3               ' 設定紙張大小
            .Orientation = Orientation.Landscape    ' 設定紙張方向
            .TopMargin = 35.941                     ' 設定上邊界(單位：px, 1cm := 28.3px, 35.941px := 1.27cm)
            .RightMargin = 35.941                   ' 設定右邊界(單位：px, 1cm := 28.3px, 35.941px := 1.27cm)
            .BottomMargin = 35.941                  ' 設定下邊界(單位：px, 1cm := 28.3px, 35.941px := 1.27cm)
            .LeftMargin = 35.941                    ' 設定左邊界(單位：px, 1cm := 28.3px, 35.941px := 1.27cm)

            ' 設定首頁、偶數頁、奇數頁頁首頁尾是否不同，文件最上層至頁首的距離、以及文件最下層至頁尾的距離
            .DifferentFirstPageHeaderFooter = False
            .HeaderDistance = 20
            .FooterDistance = 10
        End With
    End Sub

    ''' <summary>
    ''' 建立 Doc 頁首及頁尾
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteDocHeaderAndFooter()
        ' How-to: Create Headers/Footers using DocumentBuilder
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/howto-create-headersfooters-using-documentbuilder.html

        ' 1. 設定頁首，以及文字對齊方式
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary)
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center

        ' 2. 設定頁首字型、粗體以及字型大小
        With builder.Font
            .Name = "微軟正黑體"
            .Bold = True
            .Size = 18
        End With

        ' 3. 寫入文字至頁首
        builder.Write("Aspose.Words 頁首頁尾測試：這是封面頁頁首")

        ' 4. 設定頁尾，以及文字對齊方式
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary)
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center

        ' 5. 設定頁尾字型、粗體以及字型大小
        With builder.Font
            .Name = "微軟正黑體"
            .Bold = False
            .Size = 8
        End With

        ' 6. 寫入文字至頁尾
        ' Supported Fields
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/field-update.html
        builder.Write("第")
        builder.InsertField("PAGE")
        builder.Write("頁/共")
        builder.InsertField("NUMPAGES")
        builder.Write("頁")
    End Sub

    ''' <summary>
    ''' 建立 Doc 段落內容
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteDocParagraphs()
        ' Inserting Document Elements
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/inserting-document-elements.html

        ' 因為前面寫頁首頁尾時，有進行 builder.MoveToXXXXX 的動作(.MoveToHeaderPrimary, .MoveToFooterPrimary)，所以現在要記得把 builder 移回來
        ' 另外一個做法，把設定頁首頁尾放在最後執行，就不必執行此方法
        builder.MoveToSection(0)

        ' 1. 設定段落字型、粗體以及字型大小
        With builder.Font
            .Name = "新細明體"
            .Bold = False
            .Size = 12
        End With

        ' 2. 設定段落第一行縮排、對齊方式
        With builder.ParagraphFormat
            .FirstLineIndent = 56.6                     ' 設定第一行縮排(單位：px, 1cm := 28.3px, 56.6px := 2cm)
            .Alignment = ParagraphAlignment.Justify     ' 左右對齊
            .KeepTogether = True                        ' 設定為 true 時，會把段落文字放在同一頁
        End With

        ' 3. 寫入段落文章
        Dim article As StringBuilder = New StringBuilder()
        With article
            .Append("在會計處理作業中，立沖帳處理極為重要。一般在財務報表上所顯示的會計科目，可以稱之為控制科目(帳)(Control Account)或總帳科目，屬於彙總各相關明細科目(帳)(Subsidary Account)後的餘額。")
            .Append("譬如：業界常見的賒銷交易，當發生一筆賒銷交易時(假設客戶為Ａ)，就會計分錄而言，借方為應收帳款－Ａ客戶；貸方為銷貨收入。則「應收帳款」稱為控制科目；Ａ客戶即用來表示其明細科目。")
            .Append("因為就會計或管理的目的來說，不僅須了解交易對財務報表的影響，亦須了解此交易對明細科目的影響。")
            .Append("以前述例子來看，透過適當的立沖帳，才能了解哪一客戶欠公司多少的貨款(也才能據此編製應收帳款帳齡分析、了解個別客戶的信用狀況)。")
            .Append("所謂的「立沖帳」就是在設定一會計科目(控制科目／帳)時，亦須考慮該會計科目有無相關明細科目／帳，以便日後某科目之控制科目餘額會等於其各明細科目餘額。")
            .Append("會計上常見的這類科目，有：銀行存款、應收票據、應收帳款、存貨、固定資產...等。期初開帳所需要的資料包括：去年(或電腦帳上線前一個月)的資產負債 表與各科目結餘的餘額明細表。")
            .Append("其中，有關未兌現的應收／應付票據和應收／應付帳款、尚未清償的銀行借款、預收／預付貨款...等明細資料建立完成後，將來 在做沖銷(前述科目餘額減少)時，不僅可沖銷明細餘額，亦能結算相對應控制科目的餘額，此外，亦有助於稽核(Auditing)之進行。")
            .Append("大體而言，沖銷可分為餘額沖銷與逐筆沖銷；譬如：應收帳款的沖銷，餘額沖銷就是收到某一客戶的貨款後，直接從該客戶應收總額扣減一筆金額；逐筆沖銷就是客戶支付貨款時，必須先辨認原是哪一(幾)筆貨款，再予進行沖減應收帳款。")
            .Append("其中，一筆帳款對應一筆付款，最為單純，只要沖銷的鍵值(Key)設定的好，應可正確沖銷。")
            .Append("其他情況，在電腦程式設計上，則複雜的多。")
        End With

        builder.Writeln(article.ToString())
        builder.InsertParagraph()

        'builder.InsertBreak(BreakType.PageBreak)
        'builder.InsertBreak(BreakType.SectionBreakNewPage)

        builder.Writeln(article.ToString() + article.ToString())
        builder.InsertParagraph()
    End Sub

    ''' <summary>
    ''' 建立 Doc 文字內容
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteDocTextSeparately()
        ' Inserting Document Elements
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/inserting-document-elements.html

        ' 1. 設定段落第一行縮排、對齊方式
        With builder.ParagraphFormat
            .FirstLineIndent = 0                        ' 設定第一行縮排(單位：px, 1cm := 28.3px)
            .Alignment = ParagraphAlignment.Right       ' 靠右對齊
            .KeepTogether = True                        ' 設定為 true 時，會把段落文字放在同一頁
        End With

        ' 2. 設定字型、斜體、粗體、字型大小與顏色
        With builder.Font
            .Name = "Arial"
            .Italic = True
            .Bold = True
            .Size = 16
            .Color = System.Drawing.Color.Chocolate
        End With

        builder.InsertParagraph()

        ' 3. 寫一段文字
        builder.Write("ASPOSE.WORDS")

        With builder.Font
            .Name = "Calibri"
            .Italic = False
            .Bold = False
            .Size = 12
            .Color = System.Drawing.Color.Black
        End With

        builder.Write(" makes producing DOC ")

        With builder.Font
            .Name = "Comic"
            .Bold = True
            .Underline = Underline.Dash
            .Size = 16
            .Color = System.Drawing.Color.Black
        End With

        builder.Write(" SO EASY! ")

        With builder.Font
            .Name = "微軟正黑體"
            .Bold = True
            .Underline = Underline.Double
            .Size = 18
            .Color = System.Drawing.Color.Blue
        End With

        builder.Write(" 真輕鬆！")

        builder.InsertParagraph()
    End Sub

    ''' <summary>
    ''' 建立 Doc 條列內容
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteDocList()
        ' Specifying Formatting
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/specifying-formatting.html

        ' 1. 設定段落第一行縮排、對齊方式
        With builder.ParagraphFormat
            .FirstLineIndent = 0                        ' 1cm = 28.3px
            .Alignment = ParagraphAlignment.Left        ' 靠左對齊
            .KeepTogether = True                        ' 設定為 true 時，會把段落文字放在同一頁
        End With

        ' 2. 設定字型、斜體、粗體、字型大小與顏色
        With builder.Font
            .Name = "新細明體"
            .Italic = False
            .Bold = False
            .Underline = Underline.None
            .Size = 14
            .Color = System.Drawing.Color.Black
        End With

        builder.InsertParagraph()

        ' 3. 開始寫入一組 List
        builder.ListFormat.ApplyBulletDefault()

        builder.Writeln("台灣大學")
        builder.Writeln("逢甲大學")
        builder.Writeln("淡江大學")
        builder.ListFormat.ListIndent()     ' 增加一組縮排

        builder.Writeln("教學單位")
        builder.ListFormat.ListIndent()

        builder.Writeln("文學院")
        builder.Writeln("商學院")
        builder.Writeln("管理學院")

        builder.ListFormat.ListOutdent()    ' 減少一組縮排

        builder.Writeln("行政單位")
        builder.ListFormat.ListIndent()

        builder.Writeln("教務處")
        builder.Writeln("學務處")
        builder.Writeln("資訊中心")
        builder.ListFormat.ListIndent()

        builder.Writeln("作業管理組")
        builder.Writeln("校務資訊組")

        builder.ListFormat.ListOutdent()

        builder.Writeln("圖書館")

        ' 4. 結束寫入List
        builder.ListFormat.RemoveNumbers()

        builder.InsertParagraph()
    End Sub

    ''' <summary>
    ''' 建立 Doc 表格內容
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub WriteDocTable()
        ' Inserting Document Elements
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/inserting-document-elements.html

        ' Specifying Formatting
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/specifying-formatting.html

        ' How-to: AutoFit a Table to Page Width
        ' http://www.aspose.com/documentation/.net-components/aspose.words-for-.net/howto-autofit-a-table-to-page-width.html

        builder.InsertParagraph()

        ' 1. 開始一個新表格
        builder.StartTable()

        ' 2. 開始第一個 Row
        ' ===> 2-1. 設定 Row 的屬性
        builder.RowFormat.Height = 20
        builder.RowFormat.Borders.LineStyle = LineStyle.Single
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.SkyBlue
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center

        ' ===> 2-2. 開始寫入資料
        builder.InsertCell()
        builder.CellFormat.Width = 150
        builder.CellFormat.VerticalMerge = CellMerge.First
        builder.Write("工作項目")

        builder.InsertCell()
        builder.CellFormat.Width = 150
        builder.CellFormat.VerticalMerge = CellMerge.First
        builder.Write("預定完成日期")

        builder.InsertCell()
        'builder.CellFormat.HorizontalMerge = CellMerge.First
        builder.CellFormat.Width = 210
        builder.Write("本月進度")

        'builder.InsertCell()
        'builder.CellFormat.HorizontalMerge = CellMerge.Previous

        'builder.InsertCell()
        'builder.CellFormat.HorizontalMerge = CellMerge.Previous

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.First
        'builder.CellFormat.HorizontalMerge = CellMerge.None
        ' 這兩個混用後導致無法合併
        builder.CellFormat.Width = 100
        builder.Write("備註")

        ' ===> 2-3. 結束第一個 Row
        builder.EndRow()

        ' 3. 開始第二個 Row
        ' ===> 3-1. 設定 Row 的屬性
        builder.RowFormat.Height = 20
        builder.RowFormat.Borders.LineStyle = LineStyle.Single
        builder.RowFormat.Borders.Bottom.LineStyle = LineStyle.Double
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.SkyBlue
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center

        ' ===> 3-2. 開始寫入資料
        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.Previous
        builder.CellFormat.Width = 150

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.Previous
        builder.CellFormat.Width = 150

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("預定")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("實際")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("時數")

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.Previous
        builder.CellFormat.Width = 100

        ' ===> 3-3. 結束第二個 Row
        builder.EndRow()

        ' 4. 開始第三個 Row
        ' ===> 4-1. 設定 Row 的屬性
        builder.RowFormat.Height = 20
        builder.RowFormat.Borders.Bottom.LineStyle = LineStyle.Single
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center

        ' ===> 4-2. 開始寫入資料
        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.None
        builder.CellFormat.Width = 150
        builder.Write("(一) 討論預算系統修訂內容")

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.None
        builder.CellFormat.Width = 150
        builder.Write("99.10.21")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("0%")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("100%")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("8")

        builder.InsertCell()
        builder.CellFormat.VerticalMerge = CellMerge.None
        builder.CellFormat.Width = 100
        builder.Writeln("99.10.14 於會議室討論人事薪津與學雜費之預算編審功能")
        builder.Writeln("")
        builder.Writeln("99.10.21 於B217討論982預算編審系統應改進及強化之處")

        ' ===> 4-3. 結束第三個 Row
        builder.EndRow()

        ' 5. 開始第四個 Row
        ' ===> 5-1. 開始寫入資料
        builder.InsertCell()
        builder.CellFormat.Width = 150
        builder.Write("(二) 維護舊版預算處理系統")

        builder.InsertCell()
        builder.CellFormat.Width = 150
        builder.Writeln("(99.09.30)")
        builder.Writeln("(99.10.30)")
        builder.Writeln("99.11.02")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("100%")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("95%")

        builder.InsertCell()
        builder.CellFormat.Width = 70
        builder.Write("106")

        builder.InsertCell()
        builder.CellFormat.Width = 100
        builder.Writeln("1. 介面設計")
        builder.Writeln("2. 資料呈現與存取")
        builder.Writeln("3. 資料庫預儲程序")

        ' ===> 5-2. 結束第四個 Row
        builder.EndRow()

        builder.EndTable()
    End Sub

    ''' <summary>
    ''' 儲存 Doc 檔案
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SaveDocFile()
        Dim root_dir As String = Server.MapPath("~")
        Dim file_name As String = root_dir + "hello_aspose.docx"

        'builder.Document.Save(file_name, SaveFormat.Docx)
        builder.Document.Save(Response.OutputStream, SaveFormat.Docx)

        With Response
            .AddHeader("content-disposition", "attachment; filename=hello_aspose.docx")
            .ContentType = "application/octet-stream"
            .Flush()
            .Clear()
            .Close()
        End With
    End Sub

    Private Sub btnValidate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnValidate.Click
        Dim root_dir As String = Server.MapPath("~")
        Dim file_name As String = root_dir + "hello_aspose.docx"

        Dim doc As Document = New Document(file_name)

        lblMesg.Text = "File format is " + FileFormatUtil.DetectFileFormat(file_name).LoadFormat.ToString()
        lblMesg.Text += ", this file has " + doc.Sections.Count.ToString() + " sections."
    End Sub
End Class