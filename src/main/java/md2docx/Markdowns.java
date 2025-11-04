package md2docx;

import com.vladsch.flexmark.ast.BlockQuote;
import com.vladsch.flexmark.ast.BulletList;
import com.vladsch.flexmark.ast.Code;
import com.vladsch.flexmark.ast.Emphasis;
import com.vladsch.flexmark.ast.FencedCodeBlock;
import com.vladsch.flexmark.ast.Heading;
import com.vladsch.flexmark.ast.Image;
import com.vladsch.flexmark.ast.IndentedCodeBlock;
import com.vladsch.flexmark.ast.Link;
import com.vladsch.flexmark.ast.ListItem;
import com.vladsch.flexmark.ast.OrderedList;
import com.vladsch.flexmark.ast.Paragraph;
import com.vladsch.flexmark.ast.StrongEmphasis;
import com.vladsch.flexmark.ast.ThematicBreak;
import com.vladsch.flexmark.ext.tables.TableBlock;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.html.AttributeProvider;
import com.vladsch.flexmark.html.AttributeProviderFactory;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.html.IndependentAttributeProviderFactory;
import com.vladsch.flexmark.html.renderer.LinkResolverContext;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Document;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.NotNull;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Entities;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Collections;

/**
 * markdown 工具类
 *
 * @author ludangxin
 * @since 2025/10/14
 */
@Slf4j
public class Markdowns {

    public static MarkdownBuilder builder(InputStream inputStream, String charset) {
        String markdownContent = readMarkdownContent(inputStream, charset);
        return builder(markdownContent);
    }

    public static MarkdownBuilder builder(InputStream inputStream) {
        String markdownContent = readMarkdownContent(inputStream);
        return builder(markdownContent);
    }

    public static MarkdownBuilder builder(File file) {
        String markdownContent = readMarkdownContent(file);
        return builder(markdownContent);
    }

    public static MarkdownBuilder builder(String markdownContent) {
        return new MarkdownBuilder().content(markdownContent);
    }

    public static String readMarkdownContent(File file) {
        if (file == null || !file.exists()) {
            return "";
        }

        try {
            return readMarkdownContent(new FileReader(file));
        }
        catch (Exception e) {
            log.error("failed to read markdown content", e);
        }

        return "";
    }

    public static String readMarkdownContent(InputStream inputStream) {
        try {
            return readMarkdownContent(new InputStreamReader(inputStream));
        }
        catch (Exception e) {
            log.error("failed to read markdown content", e);
        }

        return "";
    }

    public static String readMarkdownContent(InputStream inputStream, String charset) {
        if (charset == null || charset.isEmpty()) {
            return readMarkdownContent(new InputStreamReader(inputStream));
        }

        try {
            return readMarkdownContent(new InputStreamReader(inputStream, charset));
        }
        catch (Exception e) {
            log.error("failed to read markdown content", e);
        }

        return "";
    }

    public static String readMarkdownContent(InputStreamReader inputStreamReader) {
        try (BufferedReader reader = new BufferedReader(inputStreamReader)) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line);
                sb.append(System.lineSeparator());
            }
            return sb.toString();
        }
        catch (IOException e) {
            log.error("failed to read markdown content", e);
        }

        return "";
    }

    public static class MarkdownBuilder {
        private String content;

        private MutableDataSet options;

        private AttributeProviderFactory attributeProviderFactory;

        private AttributeProvider attributeProvider;

        private MarkdownBuilder content(String content) {
            this.content = content;
            return this;
        }

        public MarkdownBuilder options(MutableDataSet options) {
            this.options = options;
            return this;
        }

        public MarkdownBuilder attributeProviderFactory(AttributeProviderFactory attributeProviderFactory) {
            this.attributeProviderFactory = attributeProviderFactory;
            return this;
        }

        public MarkdownBuilder attributeProvider(AttributeProvider attributeProvider) {
            this.attributeProvider = attributeProvider;
            return this;
        }

        public MarkdownBuilder printContent() {
            System.out.println(content);
            return this;
        }

        public boolean isMarkdown() {
            if (content == null || content.trim()
                                          .isEmpty()) {
                return false;
            }

            final Document document = this.buildDocument();

            return hasMarkdownNodes(document);
        }

        public Document buildDocument() {
            Parser parser = Parser.builder(this.getOptionsOrDefault())
                                  .build();

            return parser.parse(content);
        }

        public String buildHtmlContent() {
            return this.wrapperHtml(this.getHtmlRenderer()
                                        .render(this.buildDocument()));
        }

        public String buildRawHtmlContent() {
            return this.getHtmlRenderer()
                       .render(this.buildDocument());
        }

        public String buildRawHtmlIfMarkdown() {
            if (this.isMarkdown()) {
                return this.buildRawHtmlContent();
            }

            return content;
        }

        public String buildHtmlIfMarkdown() {
            if (this.isMarkdown()) {
                return this.buildHtmlContent();
            }

            return content;
        }

        private HtmlRenderer getHtmlRenderer() {
            final HtmlRenderer.Builder builder = HtmlRenderer.builder(getOptionsOrDefault());

            if (attributeProviderFactory != null) {
                builder.attributeProviderFactory(attributeProviderFactory);
            }

            if (attributeProviderFactory == null && attributeProvider != null) {
                final IndependentAttributeProviderFactory independentAttributeProviderFactory = new IndependentAttributeProviderFactory() {
                    @Override
                    public @NotNull AttributeProvider apply(@NotNull LinkResolverContext linkResolverContext) {
                        return attributeProvider;
                    }
                };
                builder.attributeProviderFactory(independentAttributeProviderFactory);
            }

            return builder.build();
        }

        private MutableDataSet getOptionsOrDefault() {
            if (options == null) {
                return this.defaultOptions();
            }
            else {
                return options;
            }
        }

        private MutableDataSet defaultOptions() {
            MutableDataSet options = new MutableDataSet();
            // 启用表格扩展，支持 Markdown 表格语法
            options.set(Parser.EXTENSIONS, Collections.singletonList(TablesExtension.create()));
            // 禁用跨列
            options.set(TablesExtension.COLUMN_SPANS, false);
            // 表头固定为 1 行
            options.set(TablesExtension.MIN_HEADER_ROWS, 1);
            options.set(TablesExtension.MAX_HEADER_ROWS, 1);
            // 自动补全缺失列、丢弃多余列
            options.set(TablesExtension.APPEND_MISSING_COLUMNS, true);
            options.set(TablesExtension.DISCARD_EXTRA_COLUMNS, true);
            return options;
        }

        private String wrapperHtml(String htmlContent) {
            org.jsoup.nodes.Document jsoupDoc = Jsoup.parse(htmlContent);
            jsoupDoc.outputSettings()
                    // 内容输出时遵循XML语法规则
                    .syntax(org.jsoup.nodes.Document.OutputSettings.Syntax.xml)
                    // 内容转义时遵循xhtml规范
                    .escapeMode(Entities.EscapeMode.xhtml)
                    // 禁用格式化输出
                    .prettyPrint(false);
            return jsoupDoc.html();
        }

        /**
         * 检查 AST 中是否存在 Markdown 特有节点（非纯文本段落）
         */
        private static boolean hasMarkdownNodes(Node node) {
            if (node == null) {
                return false;
            }

            // 判断当前节点是否为 Markdown 特有节点（非纯文本）
            if (isMarkdownSpecificNode(node)) {
                return true;
            }

            // 递归检查子节点
            Node child = node.getFirstChild();
            while (child != null) {
                if (hasMarkdownNodes(child)) {
                    return true;
                }
                child = child.getNext();
            }

            return false;
        }

        /**
         * 判定节点是否为 Markdown 特有节点（非纯文本段落）
         * 纯文本段落（Paragraph）且无任何格式（如链接、粗体等）则视为非 Markdown
         */
        private static boolean isMarkdownSpecificNode(Node node) {
            // 标题（# 标题）
            if (node instanceof Heading) {
                return true;
            }
            // 列表（有序/无序）
            if (node instanceof BulletList || node instanceof OrderedList) {
                return true;
            }
            // 列表项
            if (node instanceof ListItem) {
                return true;
            }
            // 链接（[文本](url)）
            if (node instanceof Link) {
                return true;
            }
            // 图片（![alt](url)）
            if (node instanceof Image) {
                return true;
            }
            // 粗体（**文本** 或 __文本__）
            if (node instanceof StrongEmphasis) {
                return true;
            }
            // 斜体（*文本* 或 _文本_）
            if (node instanceof Emphasis) {
                return true;
            }
            // 代码块（```代码```）
            if (node instanceof FencedCodeBlock || node instanceof IndentedCodeBlock) {
                return true;
            }
            // 表格（| 表头 | ... |）
            if (node instanceof TableBlock) {
                return true;
            }
            // 引用（> 引用内容）
            if (node instanceof BlockQuote) {
                return true;
            }
            // 水平线（--- 或 ***）
            if (node instanceof ThematicBreak) {
                return true;
            }

            // 段落节点需进一步检查是否包含 inline 格式（如粗体、链接等）
            if (node instanceof Paragraph) {
                return hasInlineMarkdownNodes(node);
            }

            // 其他节点（如文本节点）视为非特有
            return false;
        }

        /**
         * 检查段落中是否包含 inline 格式（如粗体、链接等）
         */
        private static boolean hasInlineMarkdownNodes(Node paragraph) {
            Node child = paragraph.getFirstChild();
            while (child != null) {
                // 若段落中包含任何 Markdown  inline 节点，则视为 Markdown
                if (child instanceof Link || child instanceof Image || child instanceof StrongEmphasis || child instanceof Emphasis || child instanceof Code) {
                    return true;
                }
                child = child.getNext();
            }
            return false;
        }
    }
}
