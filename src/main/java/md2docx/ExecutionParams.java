package md2docx;

import lombok.Builder;
import lombok.Data;
import lombok.Singular;
import org.apache.commons.compress.changes.ChangeSet;

import java.util.List;

/**
 * 数据库变更执行参数
 */
@Data
@Builder
public class ExecutionParams {
    /**
     * 变更集列表
     */
    @Singular("changeSet")  
    private final List<ChangeSet> changeSets;
    

    /**
     * 是否打印 SQL
     */
    @Builder.Default
    private final boolean printSql = false;

    private String name;

}