package md2docx;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.HashMap;
import java.util.Map;

/**
 * docs test
 *
 * @author ludangxin
 * @since 2025/11/4
 */
@Slf4j
public class DocxTest {
    private static final File TEMPLATE_FILE = new File("demo.docx");
    private static final File OUTPUT_FILE = new File("output.docx");
    private static final Map<String, Object> DATA = new HashMap<>();

    @BeforeAll
    public static void given_data() {
        DATA.put("user", "嘉文四世");
        DATA.put("summoner", "张铁牛");
        DATA.put("position", "打野");
        DATA.put("dialogue", "给我找些更强的敌人");
        final String markdownContent = "- **背景故事**：嘉文四世是德玛西亚国王嘉文三世的独生子，其母凯瑟琳女士因难产而死。嘉文在宫廷中长大，接受了良好的德玛西亚式教育，并结识了赵信，向其学习战争艺术。他与盖伦年龄相仿，结为好兄弟。嘉文曾率军前往边境对抗诺克萨斯，却因战力分散而战败，幸得希瓦娜相救。后来，德玛西亚国内搜魔人兵团搜捕魔法师引发起义，嘉文三世惨遭弑杀，嘉文四世接掌了议会，之后他登基成为德玛西亚国王。\n" + "\n" + "- **角色定位**：在游戏中，嘉文四世的定位是坦克、战士，他常常需要带头冲入敌方阵地，因此相比输出更加需要增强防御能力。\n" + "\n" + "- 技能介绍\n" + "  ：\n" + "  - **被动技能 - 战争律动**：普攻命中时，会对目标造成 8% 当前生命值的额外物理伤害，该效果作用于同一目标的冷却时间为 6 秒。\n" + "  - **一技能 - 巨龙撞击**：用长矛穿透路径上的敌人，对其造成物理伤害，并减少其护甲，持续 3 秒。若长矛触及 “德邦军旗”，嘉文四世会被引向军旗，并击飞沿途敌人 0.75 秒。\n" + "  - **二技能 - 黄金圣盾**：释放出一道帝王光环，使周围敌人减速，持续 2 秒，同时提供一个可以吸收伤害的护盾，持续 5 秒，附近每多一名敌方英雄，吸收伤害增加。\n" + "  - **三技能 - 德邦军旗**：投掷一柄军旗，对敌人造成魔法伤害，并将军旗置于原地 8 秒，使附近队友获得攻击速度加成。在 “德邦军旗” 附近再次点击施放该技能，将会朝军旗施放 “巨龙撞击”。\n" + "  - **终极技能 - 天崩地裂**：跃向敌方英雄，对目标及其附近的敌人造成物理伤害，并在目标周围形成环形障碍，持续 3.5 秒，再次点击施放可使障碍倒塌。\n" + "\n" + "- **皮肤信息**：嘉文四世拥有多款皮肤，包括孤胆英豪、暗星、福牛守护者等。";
        // markdown 2 html
        final String htmlContent = Markdowns.builder(markdownContent)
                                            .buildHtmlContent();
        DATA.put("description", htmlContent);
    }

    @Test
    public void given_template_doc_and_content_when_replace_then_complete() {
        Docs.builder(TEMPLATE_FILE).buildWord(DATA, OUTPUT_FILE);
    }

    @Test
    @SneakyThrows
    public void given_template_doc_and_content_when_replace_and_output_stream_then_complete() {
        final OutputStream fileOutputStream = Files.newOutputStream(OUTPUT_FILE.toPath());
        // 接收输出流
        Docs.builder(TEMPLATE_FILE).autoCloseStream(true).buildWord(DATA, fileOutputStream);
    }
}
