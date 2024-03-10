package org.ivandr.excel.mapper.fastexcel;

import com.google.common.reflect.TypeToken;
import lombok.*;
import org.ivandr.excel.annotations.ExcelExportObject;

import java.lang.invoke.MethodHandles;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.Optional;

@EqualsAndHashCode(onlyExplicitlyIncluded = true)
@Getter
@ToString
class FastExcelMappingNode {

    @EqualsAndHashCode.Include
    private final int id;

    //TODO map interface to normal class
    @ToString.Exclude
    private final Optional<ExcelExportObject> exportMetaInfo;

    @Setter
    private int headerWidth, headerHeight;
    @Setter
    private int treeHeight = 0;


    //TODO this implementation needs separating into 3 classes (better for OOP), instead of composition
    private final Optional<ValueGetter> valueGetter;
    private final Optional<CollectionGetter> collectionGetter;
    @NonNull
    private final Class<?> clazz;

    public FastExcelMappingNode(int id,
                                ExcelExportObject excelExportObject,
                                CollectionGetter collectionGetter,
                                ValueGetter valueGetter,
                                @NonNull
                                Class<?> clazz) {
        this.id = id;
        this.exportMetaInfo = Optional.ofNullable(excelExportObject);
        this.collectionGetter = Optional.ofNullable(collectionGetter);
        this.valueGetter = Optional.ofNullable(valueGetter);
        this.clazz = clazz;
        initHeaderSizes();
    }

    public <T> FastExcelMappingNode(int id,
                                    @NonNull ExcelExportObject excelExportObject,
                                    @NonNull CollectionGetter collectionGetter,
                                    @NonNull Class<?> clazz) {
        this(id, excelExportObject, collectionGetter,
                null, clazz);
    }

    public FastExcelMappingNode(int id,
                                @NonNull ExcelExportObject excelExportObject,
                                @NonNull ValueGetter valueGetter,
                                @NonNull Class<?> clazz) {
        this(id, excelExportObject, null,
                valueGetter, clazz);
    }

    public FastExcelMappingNode(int id,
                                @NonNull ExcelExportObject excelExportObject,
                                @NonNull Class<?> clazz) {
        this(id, excelExportObject, null,
                null, clazz);
    }

    public FastExcelMappingNode(int id, @NonNull Class<?> clazz) {
        this(id, null, null,
                null,
                clazz);
    }

    private void initHeaderSizes() {
        this.headerHeight = this.exportMetaInfo.isEmpty() ? 0 : 1;
        this.headerWidth = this.exportMetaInfo.isEmpty() ? 0 : 1;
    }

    public boolean isCollectionMapping() {
        return this.getCollectionGetter().isPresent() && this.getValueGetter().isEmpty();
    }

    @FunctionalInterface
    public interface ValueGetter {
        Object apply(Object object);
    }

    @FunctionalInterface
    public interface CollectionGetter {
        Collection<Object> apply(Object object);
    }
}
