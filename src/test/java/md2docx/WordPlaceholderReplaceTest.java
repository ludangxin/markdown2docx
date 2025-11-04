package md2docx;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * TODO
 *
 * @author ludangxin
 * @since 2025/11/4
 */
@Slf4j
public class WordPlaceholderReplaceTest {
    @Test
    public void given_template_doc_and_content_when_replace_then_replace() {
        final Configure templateEngineConfigure = Configure.builder().build();
        File templateFile = new File("demo.docx");
        File outputFile = new File("newDemo.docx");
        Map<String, Object> data = new HashMap<>();
        data.put("user", "嘉文四世");
        data.put("summoner", "张铁牛");
        data.put("position", "打野");
        data.put("dialogue", "给我找些更强的敌人");
        try (XWPFTemplate template = XWPFTemplate.compile(templateFile, templateEngineConfigure)) {
                template.render(data).writeToFile(outputFile.getAbsolutePath());
            }
            catch (IOException e) {
                log.error("failed to replace template word placeholder", e);
                throw new RuntimeException(e);
            }
    }
}
