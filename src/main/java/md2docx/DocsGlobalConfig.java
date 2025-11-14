package md2docx;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.RFonts;

import java.io.File;
import java.net.URI;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 全局（static）配置与注册点。
 */
@Slf4j
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class DocsGlobalConfig {
    /**
     * 注册本地字体和css font-family的映射关系
     *
     * @param cssFontFamily    css font family name
     * @param fontFile         本地文件
     */
    public static synchronized void registerPhysicalFontMapping(String cssFontFamily, File fontFile) {
        registerPhysicalFont(cssFontFamily, fontFile);
        registerFontMapping(cssFontFamily, cssFontFamily);
    }

    /**
     * 注册本地字体和css font-family的映射关系
     *
     * @param cssFontFamily    css font family name
     * @param docFontName      doc 的 font name
     * @param fontFile         本地文件
     */
    public static synchronized void registerPhysicalFontMapping(String cssFontFamily, String docFontName, File fontFile) {
        registerPhysicalFont(docFontName, fontFile);
        registerFontMapping(cssFontFamily, docFontName);
    }

    /**
     * 注册本地字体和css font-family的映射关系
     *
     * @param cssFontFamily    css font family name
     * @param fontAbstractPath 本地文件绝对路径
     */
    public static synchronized void registerPhysicalFontMapping(String cssFontFamily, String fontAbstractPath) {
        registerPhysicalFont(cssFontFamily, fontAbstractPath);
        registerFontMapping(cssFontFamily, cssFontFamily);
    }

    /**
     * 注册本地字体和css font-family的映射关系
     *
     * @param cssFontFamily    css font family name
     * @param docFontName      doc 的 font name
     * @param fontAbstractPath 本地文件绝对路径
     */
    public static synchronized void registerPhysicalFontMapping(String cssFontFamily, String docFontName, String fontAbstractPath) {
        registerPhysicalFont(docFontName, fontAbstractPath);
        registerFontMapping(cssFontFamily, docFontName);
    }

    /**
     * 注册一个全局的 CSS font-family -> doc 字体名 映射
     *
     * @param cssFontFamily css 中使用的 font-family 名称（例如 "my-font" 或 "Microsoft YaHei"）
     * @param docFontName   要写入 docx 的字体名称（例如 "Microsoft YaHei" / "SimSun" / 自定义已存在的字体名）
     */
    public static synchronized void registerFontMapping(String cssFontFamily, String docFontName) {
        if (cssFontFamily == null || cssFontFamily.trim().isEmpty()) {
            throw new IllegalArgumentException("cssFontFamily required");
        }
        if (docFontName == null || docFontName.trim().isEmpty()) {
            throw new IllegalArgumentException("docFontName required");
        }
        // 创建 RFonts 并注册到 XHTMLImporterImpl 的静态映射中
        try {
            RFonts rfonts = Context.getWmlObjectFactory().createRFonts();
            rfonts.setAscii(docFontName);
            rfonts.setHAnsi(docFontName);
            rfonts.setCs(docFontName);
            rfonts.setEastAsia(docFontName);
            XHTMLImporterImpl.addFontMapping(cssFontFamily, rfonts);
        } catch (Throwable t) {
            log.warn("failed to call XHTMLImporterImpl.addFontMapping for cssFont={}, docFont={}",
                    cssFontFamily, docFontName, t);
        }
    }

    /**
     * 注册本地字体
     *
     * @param fontAliasName    字体别名
     * @param fontAbstractPath 字体文件绝对路径
     */
    public static synchronized void registerPhysicalFont(String fontAliasName, String fontAbstractPath) {
        if (fontAliasName == null || fontAliasName.trim().isEmpty()) {
            throw new IllegalArgumentException("fontAliasName required");
        }
        if (fontAbstractPath == null || fontAbstractPath.trim().isEmpty()) {
            throw new IllegalArgumentException("fontAbstractPath required");
        }
        registerPhysicalFont(fontAliasName, new File(fontAbstractPath));
    }

    /**
     * 注册本地字体
     *
     * @param fontAliasName    字体别名
     * @param fontFile         字体文件
     */
    public static synchronized void registerPhysicalFont(String fontAliasName, File fontFile) {
        if (fontAliasName == null || fontAliasName.trim().isEmpty()) {
            throw new IllegalArgumentException("fontAliasName required");
        }
        if (fontFile == null) {
            throw new IllegalArgumentException("fontFile required");
        }
        registerPhysicalFont(fontAliasName, fontFile.toURI());
    }

    /**
     * 注册本地字体
     *
     * @param fontAliasName    字体别名
     * @param fontUri          字体文件uri
     */
    public static synchronized void registerPhysicalFont(String fontAliasName, URI fontUri) {
        if (fontAliasName == null || fontAliasName.trim().isEmpty()) {
            throw new IllegalArgumentException("fontAliasName required");
        }
        if (fontUri == null) {
            throw new IllegalArgumentException("fontUri required");
        }

        PhysicalFonts.addPhysicalFonts(fontAliasName, fontUri);
    }
}
