package md2docx;

import lombok.extern.slf4j.Slf4j;
import org.docx4j.XmlUtils;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Random;
import java.util.concurrent.ThreadLocalRandom;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.regex.PatternSyntaxException;

/**
 * doc util
 *
 * @author ludangxin
 * @since 2025/10/14
 */
@Slf4j
public class DocUtils {
    // 随机数生成器
    private static final Random RANDOM = ThreadLocalRandom.current();

    // 时间戳格式（精确到毫秒）
    private static final SimpleDateFormat TIME_FORMAT = new SimpleDateFormat("yyyyMMddHHmmssSSS");

    public static String extractText(Object o) {
        StringBuilder sb = new StringBuilder();

        if (o instanceof P) {
            P p = (P) o;
            p.getContent()
             .forEach(obj -> {
                 Object unwrapped = XmlUtils.unwrap(obj);

                 if (unwrapped instanceof R) {
                     R run = (R) unwrapped;
                     run.getContent()
                        .forEach(runObj -> {
                            Object inner = XmlUtils.unwrap(runObj);
                            if (inner instanceof Text) {
                                Text t = (Text) inner;
                                if (t.getValue() != null) sb.append(t.getValue());
                            }
                        });
                 }
                 // 如果还需要处理其他类型，比如 Tbl, Hyperlink, Comment 等，可继续扩展
             });
        }
        else {
            // fallback
            sb.append(XmlUtils.marshaltoString(o, true, false));
        }

        return sb.toString();
    }

    public static boolean isHtml(Object content) {
        if (!(content instanceof String)) {
            return false;
        }
        String contentStr = (String) content;

        // 2. 空字符串（trim后）直接返回false
        String trimmed = contentStr.trim();
        if (trimmed.isEmpty()) {
            return false;
        }

        // 规则1：匹配任意HTML标签（包括含换行符的情况）
        boolean hasTags = trimmed.matches("(?s).*<\\s*[^>]+\\s*[/]?\\s*>.*");

        // 规则2：匹配常见HTML根标签或块级标签（不区分大小写，支持换行符）
        boolean hasHtmlRootTags = trimmed.matches("(?si).*(<html>|<body>|<head>|<p>|<div>|<h\\d+>).*");

        return hasTags || hasHtmlRootTags;
    }

    /**
     * 根据动态前缀(pre)和后缀(suf)生成正则，判断模板字符串是否匹配 key
     *
     * @param templateStr 模板中提取的字符串（如"{{aaa}}"）
     * @param dataKey     数据中的key（如"aaa"）
     * @param pre         动态前缀（如"{{"）
     * @param suf         动态后缀（如"}}"）
     * @return 匹配返回true，否则返回false
     */
    public static boolean matchPlaceHolder(String templateStr, String dataKey, String pre, String suf) {
        // 输入校验
        if (templateStr == null || dataKey == null || pre == null || suf == null || dataKey.trim()
                                                                                           .isEmpty() || pre.isEmpty() || suf.isEmpty()) {
            return false;
        }

        try {
            // 1. 对前缀和后缀进行正则转义（处理特殊字符，如$、{、[等）
            String escapedPre = Pattern.quote(pre);
            String escapedSuf = Pattern.quote(suf);

            // 2. 动态生成正则表达式：^前缀\s*(.+?)\s*后缀$
            // 解释：^和$锚定整个字符串；\s*匹配key前后的空格；(.+?)捕获key
            String regex = "^" + escapedPre + "\\s*(.+?)\\s*" + escapedSuf + "$";
            Pattern pattern = Pattern.compile(regex);

            // 3. 匹配并提取key
            Matcher matcher = pattern.matcher(templateStr.trim());
            if (matcher.matches()) {
                String keyInTemplate = matcher.group(1); // 提取捕获的key
                return keyInTemplate.equals(dataKey.trim()); // 比较key
            }
        } catch (PatternSyntaxException e) {
            // 正则语法错误（如pre/suf为空，但已在输入校验中过滤）
            e.printStackTrace();
        }

        return false;
    }

    public static File createTempDocFile() {
        Path tempFile = null;
        try {
            tempFile = Files.createTempFile(generateTempFileName("temp-", null), ".docx");
        } catch (IOException e) {
            throw new RuntimeException("create temp doc file error", e);
        }
        File file = tempFile.toFile();
        // 标记临时文件在JVM退出时删除（即使发生异常也会删除）
        file.deleteOnExit();

        return file;
    }

    /**
     * 读取文件内容并写入输出流，自动关闭所有资源
     * @param file 待读取的文件
     * @param outputStream 目标输出流
     */
    public static void writeAndDeleteFile(File file, OutputStream outputStream) {
        try (BufferedInputStream reader = new BufferedInputStream(Files.newInputStream(file.toPath())))
        {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = reader.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
            outputStream.flush();
        } catch (Exception e) {
            log.error("failed to build word file", e);
        }
        finally {
            try {
                Files.deleteIfExists(file.toPath());
            }
            catch (IOException ignored) {
            }
        }
    }

    /**
     * 生成临时文件名：时间戳 + 随机数 + 可选后缀
     * 格式：yyyyMMddHHmmssSSS + 3位随机数 + .后缀（如20240520153022123456.txt）
     * @param prefix 文件名前缀（如"temp-"，若为null则无后缀）
     * @param suffix 文件名后缀（如".docx"，若为null则无后缀）
     * @return 唯一临时文件名
     */
    public static String generateTempFileName(String prefix, String suffix) {
        // 1. 生成当前毫秒级时间戳（精确到毫秒，确保时间维度唯一）
        String timestamp = TIME_FORMAT.format(new Date());

        // 2. 生成3位随机数（000-999），进一步降低同一毫秒内的冲突概率
        int randomNum = RANDOM.nextInt(1000); // 0-999的随机数
        String randomStr = String.format("%03d", randomNum); // 补0成3位字符串（如5→"005"）

        // 3. 拼接文件名（时间戳 + 随机数 + 后缀）
        StringBuilder fileName = new StringBuilder();
        fileName.append(prefix)
                .append(timestamp)
                .append(randomStr);

        // 4. 添加后缀（若不为null）
        if (suffix != null && !suffix.isEmpty()) {
            // 确保后缀以"."开头（如传入"txt"自动转为".txt"）
            if (!suffix.startsWith(".")) {
                fileName.append(".");
            }
            fileName.append(suffix);
        }

        return fileName.toString();
    }

      public static String addHtmlStyles(String html, String newStyles) {
        Document doc = Jsoup.parse(html);
        doc.outputSettings()
                    .syntax(org.jsoup.nodes.Document.OutputSettings.Syntax.xml)
                    .escapeMode(Entities.EscapeMode.xhtml)
                    .prettyPrint(false);
        Element styleElement = doc.selectFirst("style");

        if (styleElement != null) {
            styleElement.append("\n" + newStyles);
        } else {
            // 创建新的 style 标签
            Element newStyle = doc.createElement("style");
            newStyle.text(newStyles);
            Element head = doc.head();
            head.appendChild(newStyle);
        }
        return doc.html();
    }
}
