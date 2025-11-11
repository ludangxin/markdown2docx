package md2docx;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.Docx4J;
import org.docx4j.convert.in.xhtml.FormattingOption;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FontTablePart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Fonts;
import org.docx4j.wml.Style;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.List;

/**
 * TODO
 *
 * @author ludangxin
 * @since 2025/11/4
 */
@Slf4j
public class DocxStyleTest {
    private static WordprocessingMLPackage wordMLPackage;

    @BeforeAll
    @SneakyThrows
    public static void init_mainDocumentPart() {
        File templateFile = new File("output2.docx");
        wordMLPackage = WordprocessingMLPackage.load(templateFile);
    }

    @Test
    @SneakyThrows
    public void given_doc_template_when_extract_style_then_return_style_list() {
        final StyleDefinitionsPart sdp = wordMLPackage.getMainDocumentPart()
                                                      .getStyleDefinitionsPart();
        List<Style> styles = sdp.getContents()
                                .getStyle();
        log.info("docx styles length: {}", styles.size());
        for (Style style : styles) {
            String styleId = style.getStyleId();
            String name = style.getName()
                               .getVal();
            final String type = style.getType();
            log.info("styleId: {}, name: {}, type: {}", styleId, name, type);
        }
    }

    @Test
    @SneakyThrows
    public void given_doc_template_and_class_when_mapping_custom_style_then_render_doc() {
        final String html = "<html><head><style>table{border-collapse:collapse;border-spacing:0;width:100%;margin:1em 0;background-color:transparent}table th{background-color:#f7f7f7;border:1px solid#ddd;padding:8px 12px;text-align:left}table td{border:1px solid#ddd;padding:8px 12px}</style></head><body><h2 class=\"1\">嘉文四世</h2><blockquote><p class=\"customBodyText\">德玛西亚</p></blockquote><p class=\"customBodyText\"><strong>给我找些更强的敌人！</strong></p><table><thead><tr><th>列1</th><th>列2</th></tr></thead><tbody><tr><td>数据1</td><td>数据2</td></tr></tbody></table></body></html>";
        final MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage);
        // CLASS_TO_STYLE_ONLY：只认 class，不管 style 和 <strong>/<em> 等标签，相当于「纯 CSS 类驱动样式」
        // CLASS_PLUS_OTHER：class 是基础样式，style 和内联标签是补充 / 覆盖，相当于「类样式 + 局部微调样式」
        // IGNORE_CLASS: 忽略class样式
        importer.setParagraphFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
        importer.setRunFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
        importer.setTableFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
        // html转ooxml
        final List<Object> docxContent = importer.convert(html, null);
        final List<Object> docxOldContent = mainDocumentPart.getContent();
        // 清空模板内容 并 添加新的内容
        docxOldContent.clear();
        docxOldContent.addAll(docxContent);
        Docx4J.save(wordMLPackage, new File("newDemo.docx"), Docx4J.FLAG_NONE);
    }

    @Test
    public void when_generate_docx_then_complete() {
        final String html = "<html><head><style>table{border-collapse:collapse;border-spacing:0;width:100%;margin:1em 0;background-color:transparent}table th{background-color:#f7f7f7;border:1px solid#ddd;padding:8px 12px;text-align:left}table td{border:1px solid#ddd;padding:8px 12px}</style></head><body><h2 class=\"1\">嘉文四世</h2><blockquote><p class=\"customBodyText\">德玛西亚</p></blockquote><p class=\"customBodyText\"><strong>给我找些更强的敌人！</strong></p><table><thead><tr><th>列1</th><th>列2</th></tr></thead><tbody><tr><td>数据1</td><td>数据2</td></tr></tbody></table></body></html>";
        Docs.builder()
            .autoCloseStream(true)
            .buildWord(html, new File("demo.docx"));
    }

    @Test
    public void given_doc_when_get_fonts_then_print_list() {
        FontTablePart fontTablePart = wordMLPackage.getMainDocumentPart()
                                                   .getFontTablePart();

        Fonts fonts = fontTablePart.getJaxbElement();
        final List<Fonts.Font> font1 = fonts.getFont();

        for (Fonts.Font font : font1) {
            String fontName = font.getName();
            String fontFamily = font.getFamily() != null ? font.getFamily()
                                                               .getVal() : null;
            log.info("font name:{}, font family:{}", fontName, fontFamily);
        }
    }

    @Test
    public void given_doc_template_and_html_when_generate_docx_then_complete() {
        String html = "<p style=\"text-indent: 32pt;\"><span style=\"font-family: 黑体;\">三、相关公司基本情况</span></p>\n" + "<p style=\"text-indent: 32pt;\"><span style=\"font-family: 楷体;\">(一)xxx</span></p>";

        Docs.builder()
            .fontMapping("黑体", "sans-serif")
            .fontMapping("楷体", "楷体")
            .buildWord(html, new File("output3.docx"));
    }
}
