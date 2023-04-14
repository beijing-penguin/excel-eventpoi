package com.dc.eventpoi.core;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringParser {
	
    public static void main(String[] args) {
        String str1 = "H(\"Wor \" ,list .name,\"123\",456)";
        
        Map<String, Object> expMap = new HashMap<>();
        expMap.put("list.name", "姓名");
        
        
        System.out.println(parseParam(str1,expMap));
    }

    
    
    public static List<Object> parseParam(String keyStr,Map<String, Object> expMap) {
    	List<Object> paramValueList = new ArrayList<>();
    	//获取参数值
		String temp = "";
		boolean parse_quotes = false;
		boolean parse_kuohao_start = false;
		for (int i = 0; i < keyStr.length(); i++) {
			char cc = keyStr.charAt(i);
			if(cc == '(') {
				parse_kuohao_start = true;
			}else {
				if(parse_kuohao_start) {
					if(cc == '"' && parse_quotes == false){
						parse_quotes = true;
					}else {
						if(parse_quotes == true) {
							if(cc == '\"') {
								if(keyStr.charAt(i-1) == '\\') {
									temp = temp + String.valueOf(cc);
								}else {//结束
									parse_quotes = false;
								}
							}else {
								temp = temp + String.valueOf(cc);
							}
						}else {
							if(cc == ',' || cc == ')') {//结束
								int index = i-1;
								while(keyStr.charAt(index) == ' ') {
									index = index -1;
								}
								if(keyStr.charAt(index) != '"' && temp.trim().startsWith("list.") || temp.trim().startsWith("tab.")) {
									paramValueList.add(expMap.get(temp));
								}else {
									paramValueList.add(temp);
								}
								temp = "";
							}else {
								if(cc != ' ') {
									temp = temp + String.valueOf(cc);
								}
							}
						}
					}
				}
			}
		}
		return paramValueList;
    }
    
    public static String unescape(String str) {
        Pattern pattern = Pattern.compile("\\\\([\"'\\\\])");
        Matcher matcher = pattern.matcher(str);
        StringBuffer sb = new StringBuffer();

        while (matcher.find()) {
            char c = matcher.group(1).charAt(0);

            if (c == 'n') {
                matcher.appendReplacement(sb, "\n");
            } else if (c == 't') {
                matcher.appendReplacement(sb, "\t");
            } else if (c == 'r') {
                matcher.appendReplacement(sb, "\r");
            } else {
                matcher.appendReplacement(sb, "" + c);
            }
        }

        matcher.appendTail(sb);
        return sb.toString();
    }
    
    public static List<String> extractQuotedStrings(String input) {
        List<String> results = new ArrayList<>();
        Pattern pattern = Pattern.compile("\"((?:\\\\\"|[^\"])+)\"");
        Matcher matcher = pattern.matcher(input);

        while (matcher.find()) {
            String quotedString = matcher.group(1);
            String unescapedString = quotedString.replaceAll("\\\\\"", "\"");
            results.add(unescapedString);
        }

        return results;
    }
}