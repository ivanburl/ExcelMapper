package org.ivandr.excel.mapper.fastexcel;

import com.google.common.graph.Graph;
import com.google.common.graph.ImmutableGraph;
import com.google.common.graph.Traverser;
import lombok.Getter;
import lombok.NonNull;
import lombok.SneakyThrows;
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
     * leaves are in the same order as in tree (insertion order)
     */
    @NonNull
    private final List<FastExcelMappingNode> leaves;

    FastExcelMapper(@NonNull FastExcelMappingNode root, @NonNull Graph<FastExcelMappingNode> treeGraph) {
        this.root = root;
        this.treeGraph = ImmutableGraph.copyOf(treeGraph);

        if (!treeGraph.nodes().contains(root))
            throw new IllegalArgumentException("Root is not the part of graph");

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
                throw new IllegalStateException("Graph must be tree");
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
    public void mapToExcelSheet(@NonNull Worksheet worksheet, int startRow, int startColumn, T object) {
        mapHeadersToExcelSheet(startRow, startColumn, worksheet);
        mapValuesToExcelSheet(worksheet, startRow, startColumn, object);
    }

    private void mapHeadersToExcelSheet(
            int startRow, int startColumn,
            @NonNull Worksheet worksheet) {
        var nodesWithCoordinates = new ArrayDeque<FastExcelNodeWithCoordinates>();
        nodesWithCoordinates.add(new FastExcelNodeWithCoordinates(
                        this.root,
                        new ExcelCellCoordinates(startRow, startColumn)
                )
        );

        while (!nodesWithCoordinates.isEmpty()) {
            var nodeWithCoordinate = nodesWithCoordinates.pop();
            var parent = nodeWithCoordinate.node();
            var parentCoordinates = nodeWithCoordinate.coordinates();

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
            T value) {

        // index - treeHeight,
        // value - accumulatedMaximalCellNumber (the cell number required for specified tree height
        // including all children of current height)
        List<Integer> accumulatedMaximalCellNumberByTreeHeight = getAccumulatedMaximalCellNumberByHeight(value);

        HashMap<FastExcelMappingNode, List<Object>> exportObjectsByLeaf = getObjectsToExport(
                accumulatedMaximalCellNumberByTreeHeight,
                value
        );

        int maxTreeHeight = accumulatedMaximalCellNumberByTreeHeight.size();
        for (int i = 0; i < this.leaves.size(); i++) {
            var leaf = this.leaves.get(i);
            var exportMetaInfo = leaf.getExportMetaInfo()
                    .orElseThrow(() -> new IllegalStateException("Node without meta info, could not continue"));


            var listOfObjects = exportObjectsByLeaf.getOrDefault(this.leaves.get(i), new ArrayList<>());
            var leafTreeHeight = leaf.getTreeHeight();
            int sizePerObject = !leaf.isCollectionMapping() ?
                    accumulatedMaximalCellNumberByTreeHeight.get(leafTreeHeight) :
                    leafTreeHeight + 1 == maxTreeHeight ?
                            1 :
                            accumulatedMaximalCellNumberByTreeHeight.get(leafTreeHeight + 1);
            int topRow = startRow + maxTreeHeight - 1;
            for (var o : listOfObjects) {
                int leftColumn = startColumn + i;

                worksheet.value(topRow, leftColumn, o == null ? exportMetaInfo.valueFallback() : o.toString());
                var styleSetter = worksheet.range(topRow, leftColumn,
                        topRow + sizePerObject - 1, leftColumn).style();
                applyStyleToCell(styleSetter, exportMetaInfo.valueStyle());
                styleSetter.set();
                styleSetter.merge().set();
                topRow += sizePerObject;
            }
        }

    }

    private List<Integer> getAccumulatedMaximalCellNumberByHeight(T value) {
        HashMap<Integer, Integer> cellNumberByHeight = new HashMap<>();

        var stack = new ArrayDeque<FastExcelNodeWithValue>();
        stack.add(new FastExcelNodeWithValue(
                this.root,
                value
        ));

        int maxTreeHeight = 0;
        while (!stack.isEmpty()) {
            var parent = stack.pop();
            maxTreeHeight = Math.max(parent.node().getTreeHeight(), maxTreeHeight);

            var children = this.treeGraph.successors(parent.node());

            stack.addAll(
                    children
                            .stream()
                            .flatMap(n -> {
                                Object gotValue = getValueFromNode(n, parent.value());

                                Stream<?> resStream = Stream.of(gotValue);
                                if (n.isCollectionMapping() && gotValue instanceof Collection<?> coll) {
                                    cellNumberByHeight.compute(n.getTreeHeight(),
                                            (k, v) -> Math.max(1, v == null ? coll.size() : Math.max(v, coll.size())));
                                    resStream = coll.stream();
                                }
                                return resStream.map(o -> new FastExcelNodeWithValue(n, o));
                            }).toList()
            );
        }

        var res = new ArrayList<Integer>(maxTreeHeight + 1);
        for (int i = 0; i <= maxTreeHeight; i++) {
            res.add(cellNumberByHeight.getOrDefault(i, 1));
        }

        for (int i = res.size() - 2; i >= 0; i--) {
            res.set(i, res.get(i + 1) * res.get(i));
        }

        return res;
    }

    private HashMap<FastExcelMappingNode, List<Object>> getObjectsToExport(
            @NonNull List<Integer> accumulatedMaximalCellNumberByTreeHeight,
            T value
    ) {
        HashMap<FastExcelMappingNode, List<Object>> exportObjectsByLeaf = new HashMap<>();
        for (var l : this.leaves) exportObjectsByLeaf.put(l, new ArrayList<>());

        var stack = new ArrayDeque<FastExcelNodeWithValue>();
        stack.add(new FastExcelNodeWithValue(this.root, value));

        while (!stack.isEmpty()) {
            var parent = stack.pop();
            var children = this.treeGraph.successors(parent.node());

            stack.addAll(
                    children
                            .stream()
                            .flatMap(n -> {

                                        Object gotValue = getValueFromNode(n, parent.value());

                                        Stream<?> resStream = Stream.of(gotValue);

                                        if (n.isCollectionMapping() && gotValue instanceof Collection<?> coll) {
                                            int requiredSize =
                                                    accumulatedMaximalCellNumberByTreeHeight.get(n.getTreeHeight())
                                                    / (n.getTreeHeight() + 1 == accumulatedMaximalCellNumberByTreeHeight.size() ?
                                                            1 :
                                                            accumulatedMaximalCellNumberByTreeHeight.get(n.getTreeHeight() + 1));
                                            for (int i = coll.size(); i < requiredSize; i++) {
                                                coll.add(null);
                                            }
                                            resStream = coll.stream();
                                        }
                                        return resStream.map(o -> new FastExcelNodeWithValue(n, o));
                                    }
                            ).toList()
            );


            if (children.isEmpty()) {
                exportObjectsByLeaf.compute(parent.node(),
                        (k, v) -> {
                            if (v == null) v = new ArrayList<>();
                            v.add(parent.value());
                            return v;
                        });
            }
        }
        return exportObjectsByLeaf;
    }

    private Object getValueFromNode(@NonNull FastExcelMappingNode node,
                                    Object sourceValue) {
        if (sourceValue == null)
            return null;


        var valueStream = sourceValue instanceof Collection<?> coll ?
                coll.stream() : Stream.of(sourceValue);

        var nodeList = valueStream
                .map(v -> v == null ? null :
                        node.getCollectionGetter().isPresent() ?
                                node.getCollectionGetter().get().apply(v) :
                                node.getValueGetter()
                                        .orElseThrow(() ->
                                                new IllegalStateException("No getter was found in node!"))
                                        .apply(v))
                .collect(Collectors.toList());

        if (sourceValue instanceof Collection<?>) {
            return nodeList;
        }

        if (nodeList.isEmpty()) {
            throw new IllegalStateException("No value was got from getters!");
        }

        return nodeList.get(0);
    }


    private record FastExcelNodeWithValue(@NonNull FastExcelMappingNode node, @Nullable Object value) {
    }

    private record FastExcelNodeWithCoordinates(@NonNull FastExcelMappingNode node,
                                                @NonNull ExcelCellCoordinates coordinates) {
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
