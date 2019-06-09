package com.greelee.gloffice.aspose.util;

import com.aspose.cells.DateTime;
import com.google.common.collect.Lists;
import com.greelee.gloffice.aspose.constant.PatternConstant;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

import java.io.Serializable;
import java.time.LocalDateTime;
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
public class OfficeUtil implements Serializable {
    private static final long serialVersionUID = -7477548390371824035L;

    private OfficeUtil() {
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

    public static String replaceBlank(String str) {
        if (StringUtils.isNotBlank(str)) {
            return PatternConstant.SPECIAL_CHARACTERS.matcher(str).replaceAll("");
        }
        return null;
    }

    public static LocalDateTime dateTimeToLocalDateTime(DateTime cellDateTime) {
        return LocalDateTime.of(cellDateTime.getYear(), cellDateTime.getMonth(),
                cellDateTime.getDay(), cellDateTime.getHour(),
                cellDateTime.getMinute(), cellDateTime.getSecond()
        );
    }


    public static <k, v> MapModel<k, v> getMapKeyValueList(Map<k, v> map) {
        if (OfficeUtil.isNotEmpty(map)) {
            MapModel<k, v> mapModel = new MapModel<k, v>();
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
