package com.dc.eventpoi.core.enums;

import java.util.HashMap;
import java.util.Map;


/**
 * excel internal data type
 *
 * @author Jiaju Zhuang
 */
public enum CellDataTypeEnum {
    /**
     * string
     */
    STRING,
    /**
     * string
     */
    IN_LINE_STR,
    /**
     * This type of data does not need to be read in the 'sharedStrings.xml', it is only used for overuse, and the data
     * will be stored as a {@link #STRING}
     */
    DIRECT_STRING,
    /**
     * number
     */
    NUMBER,
    /**
     * boolean
     */
    BOOLEAN,
    /**
     * empty
     */
    EMPTY,
    /**
     * error
     */
    ERROR,
    /**
     * Images are currently supported only when writing
     */
    IMAGE;

    private static final Map<String, CellDataTypeEnum> TYPE_ROUTING_MAP = new HashMap<String, CellDataTypeEnum>(16);
    static {
        TYPE_ROUTING_MAP.put("s", STRING);
        TYPE_ROUTING_MAP.put("str", DIRECT_STRING);
        TYPE_ROUTING_MAP.put("inlineStr", IN_LINE_STR);
        TYPE_ROUTING_MAP.put("e", ERROR);
        TYPE_ROUTING_MAP.put("b", BOOLEAN);
        TYPE_ROUTING_MAP.put("n", NUMBER);
    }

    /**
     * Build data types
     *
     * @param cellType excel cellType
     * @return excel internal data type
     */
    public static CellDataTypeEnum buildFromCellType(String cellType) {
    	if (cellType == null || cellType.equals("")) {
            return EMPTY;
        }
        return TYPE_ROUTING_MAP.get(cellType);
    }
}
