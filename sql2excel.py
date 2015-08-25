#!/usr/bin/env python
#coding = utf-8

#
# Copyright 2015 cgwxyz
#
# Licensed under the Apache License, Version 2.0 (the "License"); you may
# not use this file except in compliance with the License. You may obtain
# a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
# WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
# License for the specific language governing permissions and limitations
# under the License.

import xlwt
import re
import sys, getopt

config = {
    "table_keylist":['CREATE', 'COLLATE', 'ENGINE', "PRIMARY KEY", "INDEX"],
    "field_keylist":['AUTO_INCREMENT', 'TINYINT', 'INT', 'BIGINT', 'SMALLINT', 'UNSIGNED', 'MEDIUMINT', 'CHAR', 'VARCHAR', 'TIME', 'DATE', 'TIMESTAMP', 'DEFAULT', 'CURRENT_TIMESTAMP', 'FLOAT', 'DECIMAL', 'DOUBLE', 'TINYBLOB', 'BLOB', 'TINYBLOB', 'TEXT', 'MEDIUMBLOB', 'MEDIUMTEXT', 'LONGTEXT', 'LONGBLOB'],
    "field_info_keylist":['AUTO_INCREMENT', 'NOT NULL', 'DEFAULT', 'COMMENT', 'UNSIGNED'],
    "re_table_name":re.compile(r"(\w+\s)*\`(\w+)\`\s\("),
    "re_field":re.compile(r'`(\w+)\`\s(\w+)(\(\d+(\,)?(\d+)?\))?(.*)\,'),
    "re_prikey":re.compile(r'[\w\s]+\(\`(\w+)\`\)(\,)?'),
    "re_index":re.compile(r'[\w\s]+\`(\w+)\`\s\((.*)\)(\,)?(\sUSING\s(\w+))?'),
    "re_pair":re.compile(r'(\w+)\=(\')?(\w+)(\')?'),
    "re_num":re.compile(r'\((\d+)\,?(\d+)?\)'),
    "re_default":re.compile(r'.*\s\'(\d+(\.\d+)?)\'.*'),
    "re_comment":re.compile(r'.*COMMENT\s\'(.*)\'(.*)(\,)?')
}

def parseFieldDesc(field, str):
    if str.find('AUTO_INCREMENT') != -1:
        field['auto_incr'] = 1

    if str.find('NOT NULL') != -1:
        field['notnull'] = 1

    if str.find('UNSIGNED') != -1:
        field['unsigned'] = 1

    if str.find('DEFAULT') != -1:
        match = config['re_default'].match(str)
        if match:
            field['default'] = match.group(1)
    
    if str.find('COMMENT') != -1:
        match = config['re_comment'].match(str)
        if match:
            field['comment'] = match.group(1)


def parseField(table_info, key, field): 
    match = config['re_field'].match(field) 
    if match:
        if not table_info.has_key('fields'):
            table_info['fields'] = {}
        if not table_info['fields'].has_key(match.group(1)):
            table_info['fields'][match.group(1)] = {}

        #字段类型
        table_info['fields'][match.group(1)]['type'] = match.group(2)
        # 字段长度
        if match.group(3):
            tmp_match = config['re_num'].match(match.group(3))
            if tmp_match:
                if tmp_match.group(2): # decimal 10,2
                    table_info['fields'][match.group(1)]['length'] = "{},{}".format(tmp_match.group(1), tmp_match.group(2))
                else:
                    table_info['fields'][match.group(1)]['length'] = tmp_match.group(1)
        #其他字段
        if match.group(6):
            parseFieldDesc(table_info['fields'][match.group(1)], match.group(6))

    return
    
def parseTableInfo(table_info, key, field):
    if key == "CREATE":
        match = config['re_table_name'].match(field) 
        if match:
            table_info['table_name'] = match.group(2)
    elif key == "PRIMARY KEY":
        match = config['re_prikey'].match(field) 
        if match:
            table_info['pri_key'] = match.group(1)
    elif key == "COLLATE":
        match = config['re_pair'].match(field)
        if match:
            table_info['COLLATE'] = match.group(3)
    elif key == "ENGINE":
        match = config['re_pair'].match(field)
        if match:
            table_info['engine'] = match.group(3)
    elif key == "INDEX":
        match = config['re_index'].match(field)
        if match:
            if not table_info.has_key('index'):
                table_info['index'] = {}
            if match.group(5):#exist index type
                table_info['index'][match.group(1)] = {'key':match.group(2), 'unique':match.group(5)}
            else:
                table_info['index'][match.group(1)] = match.group(2)
    #return 

def getSql(incoming_file):
    fp = open(incoming_file, 'r')
    lines = [line.strip() for line in  fp.readlines()]
    
    table_list = []
    table_info = {}
    
    for line in lines:
        for key in config["table_keylist"]:
            if line.find(key) != -1:
                parseTableInfo(table_info, key, line)
                break
        
        for key in config["field_keylist"]:
            if line.find(key) != -1:
                parseField(table_info, key, line)
                break

        if line.find(';') != -1:#get a new table
            table_list.append(table_info)
            table_info = {}
            
    return table_list

def write2excel(table_list, target_file):
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    
    for table_info in table_list:
        print table_info['table_name']
        sheet = book.add_sheet(table_info['table_name'], cell_overwrite_ok=True)
        i=0
        #开始写入表信息
        for val in table_info:
            if not val == 'fields' and not val == 'index' and not val == 'pri_key':
                if val == "engine":
                    sheet.write(i, 0, "表引擎")
                    sheet.write(i, 1, table_info[val])
                elif val == "table_name":
                    sheet.write(i, 0, "数据库表名")
                    sheet.write(i, 1, table_info[val])
                elif val == "COLLATE":
                    sheet.write(i, 0, "校对集")
                    sheet.write(i, 1, table_info[val])
                i = i+1

        #开始写入字段表头
        field_head_list = ["字段名", "类型长度", "自增长", "主键", "允许为空", "默认值", "注释"]
        for index, field in enumerate(field_head_list):
            sheet.write(i, index, field)
        #开始写入字段
        for val in table_info:
                if val == 'fields':
                    for (k, field_info) in table_info[val].items():
                        i += 1
                        sheet.write(i, 0, k)
                        if field_info.has_key('length'):
                            sheet.write(i, 1, "{} ({})".format(field_info['type'], field_info['length']))
                        else:
                            sheet.write(i, 1, "{}".format(field_info['type']))

                        sheet.write(i, 2, field_info['auto_incr'] if field_info.has_key('auto_incr') else '')
                        sheet.write(i, 3, 1 if k == table_info['pri_key'] else 0)
                        sheet.write(i, 4, field_info['not_null']  if field_info.has_key('not_null') else '')
                        sheet.write(i, 5, field_info['default'] if field_info.has_key('default') else '')
                        sheet.write(i, 6, field_info['comment'] if field_info.has_key('comment') else '')
        
    #end ,goto write to file
    book.save(target_file)
    return True

def usage():
    print #***.py -i new.sql -t excel.xls

def main():
    opts, args = getopt.getopt(sys.argv[1:], "hi:t:")

    incoming_file = ""
    target_file = ""
    for op, value in opts:
        if op == "-i":
            incoming_file = value
        elif op == "-t":
            target_file = value
        elif op == "-h":
            usage()
            sys.exit()
    if len(incoming_file) == 0:
        print "incoming sql file is need"
        sys.exit()
    
    if len(target_file) == 0:
        print "target excel file is need"
        sys.exit()
    
    write2excel(getSql(incoming_file), target_file)
    print 'done'

if __name__ == "__main__":
    main()
