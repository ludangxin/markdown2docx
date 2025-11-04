package md2docx;

import lombok.SneakyThrows;
import org.docx4j.Docx4J;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;

import java.io.File;

/**
 * html 2 docx
 *
 * @author ludangxin
 * @since 2025/10/14
 */
public class HtmlToDocx {
    @SneakyThrows
    public static void convertHtmlToDocx(String htmlContent, String outputFilePath) {
        // 创建 Word 文档包
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        // 设置 XHTML 导入器
        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
        // 将 HTML 内容导入到 Word 文档中
        Body body = mainDocumentPart.getJaxbElement().getBody();
        body.getContent().addAll(XHTMLImporter.convert(htmlContent, null));
        // 保存 Word 文档
        Docx4J.save(wordMLPackage, new File(outputFilePath), Docx4J.FLAG_NONE);
    }

    public static void main(String[] args) {
        String html = "<html><head><style>table{border-collapse:collapse;border-spacing:0;width:100%;margin:1em 0;background-color:transparent;}table th{background-color:#f7f7f7;border:1px solid #ddd;padding:8px 12px;text-align:left}table td{border:1px solid #ddd;padding:8px 12px}</style></head><body><h2>嘉文四世</h2>\n" + "<blockquote>\n" + "<p>德玛西亚</p>\n" + "</blockquote>\n" + "<p><strong>给我找些更强的敌人！</strong></p>\n" + "<table>\n" + "<thead>\n" + "<tr><th>列1</th><th>列2</th></tr>\n" + "</thead>\n" + "<tbody>\n" + "<tr><td>数据1</td><td>数据2</td></tr>\n" + "</tbody>\n" + "</table>\n" + "</body></html>";
        System.out.println(html);
        convertHtmlToDocx(html, "demo.docx");
    }
}
