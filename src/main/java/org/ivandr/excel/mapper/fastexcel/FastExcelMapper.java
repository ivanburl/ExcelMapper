package org.ivandr.excel.mapper.fastexcel;

import com.google.common.graph.Graph;
import com.google.common.graph.ImmutableGraph;
import com.google.common.graph.Traverser;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import org.dhatim.fastexcel.StyleSetter;
import org.dhatim.fastexcel.Worksheet;
import org.ivandr.excel.annotations.ExcelCellStyle;
import org.ivandr.excel.basics.ExcelCellCoordinates;
import org.ivandr.excel.mapper.ExcelMapper;

import javax.annotation.Nullable;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

@SuppressWarnings("UnstableApiUsage")
@Getter
public class FastExcelMapper<T> implements ExcelMapper<Worksheet, T> {

    /**
     * root of the tree
     * root does not have any sense, it is used as a way to convert forest to tree
     */
    @NonNull
    private final FastExcelMappingNode root;
    /**
     * tree representation of mapping routes
     */
    @NonNull
    private final ImmutableGraph<FastExcelMappingNode> treeGraph;
    /**
     * leaves are in the same order as in tree (insertation order)
     */
    @NonNull
    private final List<FastExcelMappingNode> leaves;

    FastExcelMapper(@NonNull FastExcelMappingNode root, @NonNull Graph<FastExcelMappingNode> treeGraph) {
        this.root = root;
        this.treeGraph = ImmutableGraph.copyOf(treeGraph);

        if (!treeGraph.nodes().contains(root))
            throw new IllegalStateException("Root is not the part of graph");

        this.leaves = this.treeGraph.nodes().stream()
                .filter(n -> this.treeGraph.successors(n).isEmpty())
                .toList();

        fitNodesSizeToRectangle();
    }


    /**
     * each node is representing header,
     * header is fitted to create full rectangle
     */
    private void fitNodesSizeToRectangle() {
        fitNodeSizeByWidth();
        fitNodeSizeByHeight();
    }

    private void fitNodeSizeByWidth() {
        var nodes = StreamSupport.stream(
                Traverser.forTree(treeGraph).breadthFirst(this.root).spliterator(),
                false).collect(Collectors.toList());
        Collections.reverse(nodes);

        for (var node : nodes) {
            var children = treeGraph.successors(node);
            node.setHeaderWidth(Math.max(1,
                            children.stream()
                                    .map(FastExcelMappingNode::getHeaderWidth)
                                    .reduce(0, Integer::sum)
                    )
            );
        }
    }

    private void fitNodeSizeByHeight() {
        var nodes = Traverser.forTree(this.treeGraph)
                .depthFirstPreOrder(this.root);
        this.root.setTreeHeight(0);

        for (var node : nodes) {
            var parent = this.treeGraph.predecessors(node);
            if (parent.isEmpty()) continue;
            if (parent.size() != 1) {
                throw new IllegalStateException("Graph must be tree, " +
                        "but for number of parents for node " +
                        "is bigger than 1");
            }
            node.setTreeHeight(parent.stream().findFirst().get().getTreeHeight() + 1);
        }

        int maxTreeHeight = this.leaves.stream().map(FastExcelMappingNode::getTreeHeight)
                .reduce(0, Integer::max);

        for (var leaf : this.leaves) {
            leaf.setHeaderHeight(1 + maxTreeHeight - leaf.getTreeHeight());
        }
    }

    @Override
    public void mapToExcelSheet(@NonNull Worksheet worksheet, int startRow, int startColumn, @NonNull T object) {
        mapHeadersToExcelSheet(startRow, startColumn, worksheet);
        mapValuesToExcelSheet(worksheet, startRow, startColumn, object);
    }

    private void mapHeadersToExcelSheet(
            int startRow, int startColumn,
            @NonNull Worksheet worksheet) {
//        var coordinates = new ArrayDeque<ExcelCellCoordinates>();
//        coordinates.add(new ExcelCellCoordinates(startRow, startColumn));
        var nodesWithCoordinates = new ArrayDeque<FastExcelNodeWithCoordinates>();
        nodesWithCoordinates.add(new FastExcelNodeWithCoordinates(
                        this.root,
                        new ExcelCellCoordinates(startRow, startColumn)
                )
        );

        while (!nodesWithCoordinates.isEmpty()) {
            var nodeWithCoordinate = nodesWithCoordinates.pop();
            var parent = nodeWithCoordinate.getNode();
            var parentCoordinates = nodeWithCoordinate.getCoordinates();

            var children = this.treeGraph.successors(parent)
                    .stream()
                    .filter(ch -> ch.getExportMetaInfo().isPresent())
                    .sorted(Comparator.comparing(ch -> ch.getExportMetaInfo().get().order()))
                    .toList();

            for (var child : children) {
                if (child.getExportMetaInfo().isEmpty()) continue;

                worksheet.value(parentCoordinates.row(), parentCoordinates.column(),
                        child.getExportMetaInfo().get().headerName());

                var styleSetter = worksheet.range(
                        parentCoordinates.row(),
                        parentCoordinates.column(),
                        parentCoordinates.row() + child.getHeaderHeight() - 1,
                        parentCoordinates.column() + child.getHeaderWidth() - 1
                ).style();
                styleSetter.merge();
                applyStyleToCell(styleSetter, child.getExportMetaInfo().get().headerStyle());

                //TODO replace with FastExcelNodeWithCoordinates
//                nodes.add(child);
                nodesWithCoordinates.add(
                        new FastExcelNodeWithCoordinates(
                                child,
                                ExcelCellCoordinates.sum(
                                        parentCoordinates,
                                        new ExcelCellCoordinates(child.getHeaderHeight(), 0)
                                )
                        )
                );

                parentCoordinates = ExcelCellCoordinates.sum(
                        parentCoordinates,
                        new ExcelCellCoordinates(0, child.getHeaderWidth())
                );
            }
        }
    }

    private void mapValuesToExcelSheet(
            @NonNull Worksheet worksheet,
            int startRow, int startColumn,
            @NonNull T value) {

        //Init column pointers with basic values
        HashMap<FastExcelMappingNode, List<FastExcelCoordinatesWithValue>> columnExportPointers =
                createColumnValueExportPointers(startRow, startColumn);

        // Produce column pointers
        var stack = new ArrayDeque<FastExcelNodeWithValue>();
        stack.add(new FastExcelNodeWithValue(
                this.root,
                value
        ));

        while (!stack.isEmpty()) {
            var parent = stack.pop();
            var children = this.treeGraph.successors(parent.getNode());
            //System.out.printf("Passed %d%n", children.size());
            //System.out.println(children);

            stack.addAll(
                    children
                            .stream()
                            .map(n -> {
                                        if (parent.getValue() == null)
                                            return new FastExcelNodeWithValue(n, null);


                                        var parentValueStream = parent.getValue() instanceof Collection<?> coll ?
                                                coll.stream() : Stream.of(parent.getValue());

                                        Stream<?> valueStream = Stream.of((Object) null);
                                        if (n.getValueGetter().isPresent()) {
                                            var getter = n.getValueGetter().get();
                                            valueStream = parentValueStream
                                                    .map(v -> {
                                                        if (v == null) return null;
                                                        return getter.apply(v);
                                                    });

                                        } else if (n.getCollectionGetter().isPresent()) {
                                            var getter = n.getCollectionGetter().get();
                                            valueStream = parentValueStream
                                                    .flatMap(v -> {
                                                        if (v == null) return Stream.of((Object) null);
                                                        var res  = getter.apply(v);
                                                        return res == null ? Stream.of((Object) null) : res.stream();
                                                    });
                                        }

                                        var res = valueStream.collect(Collectors.toList());

                                        if (res.size() <= 1 &&
                                                !n.isCollectionMapping() &&
                                                !(parent.getValue() instanceof Collection<?>)) {
                                            return new FastExcelNodeWithValue(n, res.isEmpty() ? null : res.get(0));
                                        }
                                        return new FastExcelNodeWithValue(n, res);
                                    }
                            ).toList()
            );

            if (children.isEmpty()) {
                // this is leaf so add value to appropriate column

                var cellPointers =
                        columnExportPointers.get(
                                parent.getNode()
                        );
                var lastPointer = cellPointers.get(cellPointers.size() - 1);
                var parentValue = parent.getValue();

                lastPointer.setValue(parentValue);

                int height = parentValue == null ? 1 :
                        parent.getValue() instanceof Collection<?> coll ?
                                coll.size() : 1;

                // add new pointer to export
                cellPointers.add(
                        new FastExcelCoordinatesWithValue(
                                new ExcelCellCoordinates(
                                        lastPointer.coordinates.row() + height,
                                        lastPointer.coordinates.column()),
                                null
                        )
                );
            }
        }

        //Export and fit the pointer position for table
        exportColumnPointersToWorksheet(worksheet, columnExportPointers);
    }

    private HashMap<FastExcelMappingNode, List<FastExcelCoordinatesWithValue>> createColumnValueExportPointers(
            int startRow, int startColumn
    ) {
        HashMap<FastExcelMappingNode, List<FastExcelCoordinatesWithValue>> columnExportPointers =
                new HashMap<>();

        for (int i = 0; i < this.leaves.size(); i++) {
            var leaf = this.leaves.get(i);
            var list = new ArrayList<>(List.of(
                    new FastExcelCoordinatesWithValue(
                            new ExcelCellCoordinates(
                                    startRow + leaf.getHeaderWidth() + leaf.getTreeHeight() - 1,
                                    startColumn + i),
                            null
                    )
            ));
            columnExportPointers.put(leaf, list);
        }

        return columnExportPointers;
    }

    private void exportColumnPointersToWorksheet(
            @NonNull Worksheet worksheet,
            @NonNull HashMap<FastExcelMappingNode, List<FastExcelCoordinatesWithValue>> columExportPointers) {

        var exportValues = columExportPointers.values();
        var setOfNumberOfExportValues = exportValues
                .stream()
                .map(List::size)
                .collect(Collectors.toSet());

        // validate arguments
        if (setOfNumberOfExportValues.size() != 1) {
            throw new IllegalArgumentException("The number of export values must be equal for each leaf (final column)!");
        }

        int numberOfExportValues = setOfNumberOfExportValues.stream().findFirst().get();

        // fit export values
        for (int i = 0; i < numberOfExportValues; i++) {
            int effectivelyFinalIndex = i;
            var rowMax = exportValues.stream()
                    .map(v -> v.get(effectivelyFinalIndex).getCoordinates().row())
                    .max(Integer::compare)
                    .orElseThrow(() -> new IllegalStateException("Unexpected state, nothing to export ..."));

            for (var exportValue : exportValues) {
                exportValue.get(i).setCoordinates(
                        new ExcelCellCoordinates(
                                rowMax,
                                exportValue.get(i).getCoordinates().column()
                        )
                );
            }
        }
        //export values
        var exportEntities = columExportPointers.entrySet();
        for (var e : exportEntities) {

            var mappingNode = e.getKey();
            var exportInfo = e.getValue();

            if (mappingNode.getExportMetaInfo().isEmpty()) continue;
            var exportMetaInfo = mappingNode.getExportMetaInfo().get();
            for (int i = 0; i < exportInfo.size() - 1; i++) {
                int startColumn = exportInfo.get(i).getCoordinates().column();
                int startRow = exportInfo.get(i).getCoordinates().row();
                int endRow = exportInfo.get(i + 1).getCoordinates().row();

                var exportValue = exportInfo.get(i).getValue();
                if (exportValue != null && mappingNode.isCollectionMapping()) {
                    var listValue = ((Collection<?>) exportValue).stream().toList();

                    for (int j = startRow; j < endRow; j++) {
                        var exportStringValue = (listValue.size() <= j - startRow) ? "" :
                                listValue.get(j - startRow) == null ?
                                        exportMetaInfo.valueFallback() :
                                        listValue.get(j - startRow).toString();

                        worksheet.value(j, startColumn, exportStringValue);
                        applyStyleToCell(worksheet.style(j, startColumn), exportMetaInfo.valueStyle());
                    }
                } else {
                    var exportStringValue = exportValue == null ?
                            exportMetaInfo.valueFallback() :
                            exportValue.toString();
                    worksheet.value(startRow, startColumn, exportStringValue);
                    var styleSetter = worksheet
                            .range(startRow, startColumn,
                                    endRow - 1, startColumn)
                            .style();
                    applyStyleToCell(styleSetter, exportMetaInfo.valueStyle());
                    styleSetter.merge().set();
                }
            }

        }
    }

    @AllArgsConstructor
    @Setter
    @Getter
    private static class FastExcelCoordinatesWithValue {
        @NonNull
        private ExcelCellCoordinates coordinates;
        @Nullable
        private Object value;
    }

    @AllArgsConstructor
    @Getter
    private static class FastExcelNodeWithValue {
        @NonNull
        private final FastExcelMappingNode node;
        @Nullable
        private final Object value;
    }

    @AllArgsConstructor
    @Getter
    private static class FastExcelNodeWithCoordinates {
        @NonNull
        private final FastExcelMappingNode node;
        @NonNull
        private final ExcelCellCoordinates coordinates;
    }


    //TODO move to another class (this function breaks single responsibility principle)
    private void applyStyleToCell(@NonNull StyleSetter styleSetter,
                                  @NonNull ExcelCellStyle style) {

        styleSetter = styleSetter.borderStyle(style.borderStyle())
                .fontSize(style.fontSize())
                .borderStyle(style.borderStyle())
                .horizontalAlignment(style.horizontalAlignment().getHorizontalAlignmentTag())
                .verticalAlignment(style.verticalAlignment().getVerticalAlignmentFastExcelTag())
                .wrapText(style.isWrapText());
        if (style.isBold())
            styleSetter = styleSetter.bold();
        if (style.isItalic())
            styleSetter = styleSetter.italic();
        if (style.isUnderlined())
            styleSetter = styleSetter.underlined();

        if (!style.cellFormat().isBlank() && !style.cellFormat().isEmpty())
            styleSetter = styleSetter.format(style.cellFormat());

        styleSetter.set();
    }
}
