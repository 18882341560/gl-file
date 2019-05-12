package com.greelee.glfile.aspose.util;

import com.google.common.collect.Lists;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * @author: gl
 * @Email: 110.com
 * @version: 1.0
 * @Date: 2019/5/12
 * @describe: todo 最后放入tool模块
 */
public class DocumentUtil implements Serializable {
    private static final long serialVersionUID = -7477548390371824035L;

    private DocumentUtil() {
    }


    public static String getSuffix(String path) {
        return path.substring(path.lastIndexOf(".") + 1);
    }


    public static <k, v> boolean isNotEmpty(Map<k, v> map) {
        return null != map && !map.isEmpty();
    }

    public static <E> boolean isNotEmpty(Collection<E> collection) {
        return null != collection && !collection.isEmpty();
    }

    @SuppressWarnings("unchecked")
    public static <k, v> MapModel getMapKeyValueList(Map<k, v> map) {
        if (DocumentUtil.isNotEmpty(map)) {
            MapModel mapModel = new MapModel<k, v>();
            List<k> kList = Lists.newArrayList();
            List<v> vList = Lists.newArrayList();
            map.forEach((k1, v1) -> {
                kList.add(k1);
                vList.add(v1);
            });
            mapModel.setKList(kList);
            mapModel.setVList(vList);
            return mapModel;
        }
        return null;
    }

    @Data
    public static class MapModel<k, v> {
        private List<k> kList;
        private List<v> vList;
    }
}
