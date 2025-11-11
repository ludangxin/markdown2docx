package md2docx;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.convert.in.xhtml.FormattingOption;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.RFonts;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.BiFunction;

/**
 * doc操作工具类
 *
 * @author ludangxin
 * @since 2025/10/14
 */
@Slf4j
public class Docs {
    @SneakyThrows
    public static DocBuilder builder() {
        return new DocBuilder().wordMLPackage(WordprocessingMLPackage.createPackage());
    }

    @SneakyThrows
    public static DocBuilder builder(File file) {
        return new DocBuilder().templateInputStream(Files.newInputStream(file.toPath()))
                               .wordMLPackage(WordprocessingMLPackage.load(file));
    }

    @SneakyThrows
    public static DocBuilder builder(InputStream inputStream) {
        return new DocBuilder().templateInputStream(inputStream)
                               .wordMLPackage(WordprocessingMLPackage.load(inputStream));
    }

    @SneakyThrows
    public static DocBuilder builder(String filePath) {
        return new DocBuilder().templateInputStream(Files.newInputStream(new File(filePath).toPath()))
                               .wordMLPackage(WordprocessingMLPackage.load(new File(filePath)));
    }

    public static class DocBuilder {
        private InputStream templateInputStream;

        private WordprocessingMLPackage wordMLPackage;

        private XHTMLImporterImpl importer;

        private FormattingOption paragraphFormatting;

        private FormattingOption runFormatting;

        private FormattingOption tableFormatting;

        private String staticResourceBaseUri;

        private String[] placeHolderPreSuffix = new String[]{"{{", "}}"};

        private Configure templateEngineConfigure;

        private boolean useHtmlDefaultStyle = true;

        private boolean autoCloseStream = true;

        private final Map<String, String> fontMapping = new HashMap<>();

        private String globalCss = "table{border-collapse:collapse;border-spacing:0;width:100%;margin:1em 0;background-color:transparent;}table th{background-color:#f7f7f7;border:1px solid #ddd;padding:8px 12px;text-align:left}table td{border:1px solid #ddd;padding:8px 12px}";

        /**
         * <String, String, String>: htmlContent htmlKey resultHtmlContent
         */
        private BiFunction<String, String, String> htmlContentProcessor;

        private DocBuilder templateInputStream(InputStream templateInputStream) {
            this.templateInputStream = templateInputStream;
            return this;
        }

        private DocBuilder wordMLPackage(WordprocessingMLPackage wordMLPackage) {
            this.wordMLPackage = wordMLPackage;
            return this;
        }

        public DocBuilder importer(XHTMLImporterImpl importer) {
            this.importer = importer;
            return this;
        }

        public DocBuilder paragraphFormatting(FormattingOption paragraphFormatting) {
            this.paragraphFormatting = paragraphFormatting;
            return this;
        }

        public DocBuilder runFormatting(FormattingOption runFormatting) {
            this.runFormatting = runFormatting;
            return this;
        }

        public DocBuilder tableFormatting(FormattingOption tableFormatting) {
            this.tableFormatting = tableFormatting;
            return this;
        }

        public DocBuilder useHtmlDefaultStyle(boolean useHtmlDefaultStyle) {
            this.useHtmlDefaultStyle = useHtmlDefaultStyle;
            return this;
        }

        public DocBuilder staticResourceBaseUri(String staticResourceBaseUri) {
            this.staticResourceBaseUri = staticResourceBaseUri;
            return this;
        }

        public DocBuilder placeHolderPreSuffix(String placeHolderPrefix, String placeHolderSuffix) {
            this.placeHolderPreSuffix = new String[]{placeHolderPrefix, placeHolderSuffix};
            return this;
        }

        public DocBuilder templateEngineConfigure(Configure templateEngineConfigure) {
            this.templateEngineConfigure = templateEngineConfigure;
            return this;
        }

        public DocBuilder autoCloseStream(boolean autoCloseStream) {
            this.autoCloseStream = autoCloseStream;
            return this;
        }

        public DocBuilder globalCss(String globalCss) {
            this.globalCss = globalCss;
            return this;
        }

        public DocBuilder htmlContentProcessor(BiFunction<String, String, String> htmlContentProcessor) {
            this.htmlContentProcessor = htmlContentProcessor;
            return this;
        }

        public DocBuilder fontMapping(Map<String, String> fontMapping) {
            this.fontMapping.putAll(fontMapping);
            return this;
        }

        public DocBuilder fontMapping(String cssFontFamily, String docFontName) {
            this.fontMapping.put(cssFontFamily, docFontName);
            return this;
        }

        public List<Object> buildWordML(String html) {
            return this.buildWordML(html, null);
        }

        public void buildWord(String html, String outputFile) {
            this.buildWord(html, new File(outputFile));
        }

        public void buildWord(String html, File outputFile) {
            try {
                this.getMainContent()
                    .addAll(this.buildWordML(html));
                wordMLPackage.save(outputFile);
            }
            catch (Exception e) {
                log.error("failed to build word file", e);
                throw new RuntimeException(e);
            }
        }

        public void buildWord(String html, OutputStream outputStream) {
            try {
                this.getMainContent()
                    .addAll(this.buildWordML(html));
                wordMLPackage.save(outputStream);
            }
            catch (Exception e) {
                log.error("failed to build word file", e);
                throw new RuntimeException(e);
            }
            finally {
                try {
                    if (autoCloseStream) {
                        outputStream.close();
                    }
                }
                catch (IOException ignored) {
                }
            }
        }

        public void buildWord(Map<String, Object> placeHolderData, OutputStream outputStream) {
            try {
                // 替换模板中的普通占位符
                if (this.checkPlaceHolderDataType(placeHolderData) == 1) {
                    this.replacePlaceHolder(placeHolderData, outputStream);
                }

                // 替换模板中包含的html
                if (this.checkPlaceHolderDataType(placeHolderData) == 2) {
                    this.replaceHtmlPlaceHolder(placeHolderData, outputStream);
                }

                // 替换普通/html占位符
                if (this.checkPlaceHolderDataType(placeHolderData) == 3) {
                    final File tempDocFile = DocUtils.createTempDocFile();
                    this.replaceHtmlPlaceHolder(placeHolderData, tempDocFile);
                    this.replacePlaceHolder(placeHolderData, tempDocFile, tempDocFile);
                    DocUtils.writeAndDeleteFile(tempDocFile, outputStream);
                }
            }
            catch (Exception e) {
                log.error("failed to build word file", e);
                throw new RuntimeException(e);
            }
            finally {
                try {
                    if (autoCloseStream) {
                        outputStream.close();
                    }
                }
                catch (IOException ignored) {
                }
            }
        }

        public void buildWord(Map<String, Object> placeHolderData, File outputFile) {
            // 替换模板中的普通占位符
            if (this.checkPlaceHolderDataType(placeHolderData) == 1) {
                this.replacePlaceHolder(placeHolderData, outputFile);
            }

            // 替换模板中包含的html
            if (this.checkPlaceHolderDataType(placeHolderData) > 1) {
                this.replaceHtmlPlaceHolder(placeHolderData, outputFile);
            }

            // 追加替换普通占位符
            if (this.checkPlaceHolderDataType(placeHolderData) == 3) {
                this.replacePlaceHolder(placeHolderData, outputFile, outputFile);
            }
        }

        private List<Object> buildWordML(String html, String htmlKey) {
            final XHTMLImporterImpl importer = this.getImporterOrDefault();
            try {
                if (globalCss != null && !globalCss.isEmpty()) {
                    html = DocUtils.addHtmlStyles(html, globalCss);
                }

                if (htmlContentProcessor != null) {
                    html = htmlContentProcessor.apply(html, htmlKey);
                }

                return importer.convert(html, staticResourceBaseUri);
            }
            catch (Exception e) {
                log.error("failed to convert HTML to XHTML", e);
                throw new RuntimeException(e);
            }
        }

        private void replaceHtmlPlaceHolder(Map<String, Object> placeHolderData, File outputFile) {
            this.doReplaceHtmlPlaceHolder(placeHolderData);

            try {
                // 替换html
                wordMLPackage.save(outputFile);
            }
            catch (Docx4JException e) {
                log.error("failed to build word file", e);
                throw new RuntimeException(e);
            }
        }

        private void replaceHtmlPlaceHolder(Map<String, Object> placeHolderData, OutputStream outputStream) {
            this.doReplaceHtmlPlaceHolder(placeHolderData);

            try {
                // 替换html
                wordMLPackage.save(outputStream);
            }
            catch (Docx4JException e) {
                log.error("failed to build word file", e);
                throw new RuntimeException(e);
            }
            finally {
                try {
                    if (autoCloseStream) {
                        outputStream.close();
                    }
                }
                catch (IOException ignored) {
                }
            }
        }

        private void doReplaceHtmlPlaceHolder(Map<String, Object> placeHolderData) {
            final List<Object> mainContent = this.getMainContent();
            List<Object> newContent = new ArrayList<>();

            for (Object p : mainContent) {
                String text = DocUtils.extractText(p);

                Optional<String> matchedKey = placeHolderData.keySet()
                                                             .stream()
                                                             .filter(key -> DocUtils.matchPlaceHolder(text, key, placeHolderPreSuffix[0], placeHolderPreSuffix[1]))
                                                             .findFirst();

                if (matchedKey.isPresent()) {
                    String key = matchedKey.get();
                    Object value = placeHolderData.get(key);

                    if (DocUtils.isHtml(value)) {
                        final List<Object> wordFragment = this.buildWordML((String) value, key);
                        newContent.addAll(wordFragment);
                    }
                    else {
                        newContent.add(p);
                    }
                }
                else {
                    newContent.add(p);
                }
            }

            // 替换模板内容
            mainContent.clear();
            mainContent.addAll(newContent);
        }

        private void replacePlaceHolder(Map<String, Object> data, File templateFile, File outputFile) {
            final Configure templateEngineConfigure = this.getTemplateEngineConfigureOrDefault();
            try (XWPFTemplate template = XWPFTemplate.compile(templateFile, templateEngineConfigure)){
                template.render(data)
                        .writeToFile(outputFile.getAbsolutePath());
            }
            catch (IOException e) {
                log.error("failed to replace template word placeholder", e);
                throw new RuntimeException(e);
            }
        }

        public void replacePlaceHolder(Map<String, Object> data, File outputFile) {
            final Configure templateEngineConfigure = this.getTemplateEngineConfigureOrDefault();

            if (templateInputStream == null) {
                throw new NullPointerException("template file can not be null");
            }

            XWPFTemplate template = XWPFTemplate.compile(templateInputStream, templateEngineConfigure);
            try {
                template.render(data)
                        .writeToFile(outputFile.getAbsolutePath());
            }
            catch (IOException e) {
                log.error("failed to replace template word placeholder", e);
                throw new RuntimeException(e);
            }
        }

        public void replacePlaceHolder(Map<String, Object> data, String outputFileAbsolutePath) {
            final Configure templateEngineConfigure = this.getTemplateEngineConfigureOrDefault();

            if (templateInputStream == null) {
                throw new NullPointerException("template file can not be null");
            }

            XWPFTemplate template = XWPFTemplate.compile(templateInputStream, templateEngineConfigure);
            try {
                template.render(data)
                        .writeToFile(outputFileAbsolutePath);
            }
            catch (IOException e) {
                log.error("failed to replace template word placeholder", e);
                throw new RuntimeException(e);
            }
        }

        public void replacePlaceHolder(Map<String, Object> data, OutputStream outputStream) {
            final Configure templateEngineConfigure = this.getTemplateEngineConfigureOrDefault();

            if (templateInputStream == null) {
                throw new NullPointerException("template file can not be null");
            }

            try {
                XWPFTemplate template = XWPFTemplate.compile(templateInputStream, templateEngineConfigure);
                final XWPFTemplate render = template.render(data);
                render.write(outputStream);
            }
            catch (IOException e) {
                log.error("failed to replace template word placeholder", e);
                throw new RuntimeException(e);
            }
            finally {
                try {
                    if (autoCloseStream) {
                        outputStream.close();
                    }
                }
                catch (IOException ignored) {
                }
            }
        }

        private XHTMLImporterImpl getImporterOrDefault() {
            XHTMLImporterImpl importerResult;
            if (importer == null) {
                if (paragraphFormatting != null || runFormatting != null || tableFormatting != null) {
                    XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage);
                    importer.setParagraphFormatting(paragraphFormatting == null ? FormattingOption.CLASS_PLUS_OTHER : paragraphFormatting);
                    importer.setRunFormatting(runFormatting == null ? FormattingOption.CLASS_PLUS_OTHER : runFormatting);
                    importer.setTableFormatting(tableFormatting == null ? FormattingOption.CLASS_PLUS_OTHER : tableFormatting);

                    importerResult =  importer;
                }
                else {
                    importerResult = this.defaultImporter();
                }
            }
            else {
                importerResult = this.importer;
            }

            if (!fontMapping.isEmpty()) {
                fontMapping.forEach((cssFont, docFont) -> {
                    RFonts rfonts = Context.getWmlObjectFactory()
                                           .createRFonts();
                    rfonts.setAscii(docFont);
                    rfonts.setHAnsi(docFont);
                    rfonts.setCs(docFont);
                    XHTMLImporterImpl.addFontMapping(cssFont, rfonts);
                });
            }

            return importerResult;
        }

        private Configure getTemplateEngineConfigureOrDefault() {
            if (templateEngineConfigure == null) {
                return this.defaultTemplateEngineConfigure();
            }
            else {
                return this.templateEngineConfigure;
            }
        }

        private Configure defaultTemplateEngineConfigure() {
            return Configure.builder()
                            .buildGramer(placeHolderPreSuffix[0], placeHolderPreSuffix[1])
                            .build();
        }

        private XHTMLImporterImpl defaultImporter() {
            XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage);
            if (useHtmlDefaultStyle) {
                importer.setParagraphFormatting(FormattingOption.CLASS_PLUS_OTHER);
                importer.setRunFormatting(FormattingOption.CLASS_PLUS_OTHER);
                importer.setTableFormatting(FormattingOption.CLASS_PLUS_OTHER);
            }
            else {
                importer.setParagraphFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
                importer.setRunFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
                importer.setTableFormatting(FormattingOption.CLASS_TO_STYLE_ONLY);
            }

            return importer;
        }

        private List<Object> getMainContent() {
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            if (globalCss != null && !globalCss.isEmpty()) {
                mainDocumentPart.getStyleDefinitionsPart()
                                .setCss(globalCss);
            }
            Body body = mainDocumentPart.getJaxbElement()
                                        .getBody();
            return body.getContent();
        }

        /**
         * 判断占位符数据类型
         *
         * @param placeHolderData 占位符数据
         * @return 1: 数据不包含html 2: 数据全是html 3: 都包含
         */
        private int checkPlaceHolderDataType(Map<String, Object> placeHolderData) {
            boolean hasHtmlValFlag = false;
            boolean hasCommonValFlag = false;

            for (Object value : placeHolderData.values()) {
                if (DocUtils.isHtml(value)) {
                    hasHtmlValFlag = true;
                }
                else {
                    hasCommonValFlag = true;
                }
            }

            if (!hasHtmlValFlag && hasCommonValFlag) {
                return 1;
            }

            if (hasHtmlValFlag && !hasCommonValFlag) {
                return 2;
            }

            return 3;
        }
    }
}