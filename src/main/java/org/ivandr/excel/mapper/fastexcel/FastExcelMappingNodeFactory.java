package org.ivandr.excel.mapper.fastexcel;

import lombok.NonNull;
import lombok.SneakyThrows;
import org.ivandr.excel.annotations.ExcelExportObject;

import java.lang.invoke.*;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.Collection;

class FastExcelMappingNodeFactory {
    @NonNull
    private final MethodHandles.Lookup lookup;

    private int id;

    public FastExcelMappingNodeFactory() {
        lookup = MethodHandles.lookup();
        id = 0;
    }

    /**
     * Reflect and creates lambda function from method
     * Further information and inspiration could be found on
     * <a href='https://norswap.com/fast-java-reflection/'>blog</a> and
     * <a href='https://gist.github.com/norswap/09846a75092f49a7f1cbf1f00f85e9b6'>gist git</a>
     * @param method - method to be reflected and compiled
     * @return compiled node for mapping
     */
    @SneakyThrows
    @NonNull
    public FastExcelMappingNode createFastExcelMappingNode(@NonNull Method method) {
        // creates c style call from this.getter -> f(this) (low level implementation)
        MethodHandle handle = lookup.unreflect(method);
        // get info from cell
        var metaInfo = method.getAnnotation(ExcelExportObject.class);

        if (Collection.class.isAssignableFrom(method.getReturnType())) {
            // create
            ParameterizedType ret = ((ParameterizedType) method.getGenericReturnType());
            Type param = ret.getActualTypeArguments()[0];
            // creates the lambda function from
            // 1 - bytecode implementation
            // 2 - name of function from interface
            // 3 - interface itself
            // 4 - parameters returned type + args
            // 5 - bytecode implementation
            // 6 - the type of implementation itself
            CallSite site = LambdaMetafactory.metafactory(
                    lookup, "apply",
                    MethodType.methodType(FastExcelMappingNode.CollectionGetter.class),
                    MethodType.methodType(Collection.class, Object.class),
                    handle, handle.type());
            return new FastExcelMappingNode(id++,
                    metaInfo,
                    (FastExcelMappingNode.CollectionGetter) site.getTarget().invoke(),
                    (Class<?>) param);
        }

        CallSite site = LambdaMetafactory.metafactory(
                lookup, "apply",
                MethodType.methodType(FastExcelMappingNode.ValueGetter.class),
                MethodType.methodType(Object.class, Object.class),
                handle, handle.type());
        return new FastExcelMappingNode(id++,
                metaInfo,
                (FastExcelMappingNode.ValueGetter) site.getTarget().invoke(),
                method.getReturnType());
    }

    /**
     * Must be used only for creating roots
     * @param clazz - class which it implements
     * @return node (without any lambda functions)
     */
    @NonNull
    public FastExcelMappingNode createFastExcelMappingNode(@NonNull Class<?> clazz) {
        return new FastExcelMappingNode(id++, clazz);
    }
}
