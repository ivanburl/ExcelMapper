package org.ivandr.excel.mapper.fastexcel;

import com.google.common.annotations.Beta;
import com.google.common.graph.*;
import lombok.NonNull;
import lombok.SneakyThrows;
import lombok.experimental.SuperBuilder;
import org.dhatim.fastexcel.Worksheet;
import org.ivandr.excel.annotations.ExcelExportObject;
import org.ivandr.excel.basics.ExcelCellCoordinates;
import org.ivandr.excel.mapper.ExcelMapper;
import org.ivandr.excel.mapper.ExcelMapperFactory;

import javax.print.DocFlavor;
import java.lang.reflect.Method;
import java.util.*;
import java.util.stream.Collectors;

@SuppressWarnings("UnstableApiUsage")
public class FastExcelMapperFactory implements ExcelMapperFactory<Worksheet> {
    @NonNull
    @SneakyThrows
    public <T> ExcelMapper<Worksheet, T> createExcelMapperForClass(@NonNull Class<T> clazz) {
        var nodeFactory = new FastExcelMappingNodeFactory();
        var root = nodeFactory.createFastExcelMappingNode(clazz);
        var graph = createMapperGraph(nodeFactory, root);
        return new FastExcelMapper<>(root, graph);
    }

    @NonNull
    private Graph<FastExcelMappingNode> createMapperGraph(@NonNull FastExcelMappingNodeFactory factory,
                                                          @NonNull FastExcelMappingNode root) {
        MutableGraph<FastExcelMappingNode> graph = GraphBuilder
                .directed()
                .nodeOrder(ElementOrder.sorted(Comparator.comparing(FastExcelMappingNode::getId)))
                .build();

        var stack = new ArrayDeque<FastExcelMappingNode>();
        stack.add(root);

        while (!stack.isEmpty()) {
            var parent = stack.pop();
            if (parent.getExportMetaInfo().isPresent() &&
                    !parent.getExportMetaInfo().get().isRecursive()) {
                continue;
            }

            var getters = getExportGettersFromClass(parent.getClazz());
            getters.sort(Comparator.comparing(g -> g.getAnnotation(ExcelExportObject.class).order()));

            for (var g : getters) {
                var child = factory.createFastExcelMappingNode(g);
                graph.putEdge(parent, child);
                stack.add(child);
            }
        }

        return graph;
    }


    @NonNull
    private List<Method> getExportGettersFromClass(Class<?> obejctClass) {
        return Arrays.stream(
                        obejctClass.getMethods()
                )
                .filter(m -> Objects.nonNull(m) &&
                        m.getAnnotationsByType(ExcelExportObject.class).length != 0 &&
                        isMethodGetter(m))
                .collect(Collectors.toList());
    }

    private boolean isMethodGetter(@NonNull Method m) {
        return m.getParameterCount() == 0 &&
                !m.getReturnType().equals(void.class);
    }
}
