#!/usr/bin/env python
# -*- encoding: utf-8 -*-

import docx

class Text(object):
    """文本对象"""
    def __init__(self, text):
        self.text= text


class Image(object):
    """图片对象"""
    def __init__(self, title, identity, text):
        self.title = title
        self.identity = identity
        self.text = text

class Table(object):
    """表格对象"""
    def __init__(self, title, datas, table_type):
        self.title = title
        self.datas = datas
        self.table_type = table_type


class Parts(object):
    """
    不同集合部分的父类
    """
    _TYPE = None
    def __init__(self, iterable):
        if _TYPE is None:
            raise ValueError("_TYPE is None")
        self._items = []
        for item in iterable:
            if _TYPE in ["image"]:
                item = Image(**item)
            elif _TYPE in ["table"]:
                item = Table(**item)
            elif _TYPE in ["text"]:
                item = Text(item)
            else:
                raise ValueError("_TYPE can't be %s" % _TYPE)
            self._items.append(item)

    def __getitem__(self, index):
        """重写self[i]"""
        return self._items[index]

    def __getslice__(self, start=None, end=None):
        """重写self[i:j]"""
        return self._items[start: end]

    def __getattr__(self, attr):
        """重写self.attr"""
        _rel_in_line = False # 判断元素之间的关系，是否在同一行
        # 获取所有元素
        if attr == "all":
            return (self[:], _rel_in_line)
        # 获取单个元素
        elif "-" not in attr or ":" not in attr:
            return ([self[int(attr)]], _rel_in_line)
        else:
            # 获取多个元素
            if "-" in attr: 
                start, end = attr.split("-")
            if ":" in attr: 
                start, end = attr.split(":")
                _rel_in_line = True
            if start == "":
                start = None
            if end == "":
                end = None
            if start:
                start = int(start)
            if end:
                end = int(end)
            return (self[start, end], _rel_in_line)

    def __iter__(self):
        return self._items

        
class Images(Parts):
    """图片集合"""
    _TYPE = "image"

class Texts(Parts):
    """文本集合"""
    _TYPE = "text"

class Tables(Parts):
    """表格集合"""
    _TYPE = "table"
    

class Section(object):
    """区域对象"""
    def __init__(self, title, images, texts, tables, file_manager, sequence=[]):
        """
        title: 标题
        images: 图片
        texts: 文字
        tables: 表格
        sequence: 标题，图片， 文字，表格，的排列顺序.eg:
            ["title",
            "images.1-3",
            "texts",
            "tables.0",
            "tables.1:2",
            "tables.1:"
            ]
        """
        self.title = Texts([title])
        self.images = Images(images)
        self.texts = Texts(texts)
        self.tables = Tables(tables)
        self.file_manager = file_manager
        self.sequence = sequence
        self._render_items = []

    def gen_render_items(self):
        """得到要渲染的真实数据"""
        for it in sequence:
            attr_name, pos = self._get_attr_pos(it)
            attr = getattr(self, attr_name)
            item = getattr(attr, pos)
            self._render_items.append(item)

    def _get_attr_pos(self, seq_item):
        """
        根据sequence中的元素，获取属性对象名和属性对象中真实数据的位置信息
        """
        if "." in seq_item:
            attr_pos = seq_item.split(".")
        else:
            attr_pos = seq_item
        if len(attr_pos) == 1:
            attr = attr_pos[0]
            pos = "all"
        elif len(attr_pos) > 1:
            attr, pos = attr_pos
        return attr, pos
            
    def render(self, doc):
        """渲染区域
        doc: docx.Document对象
        """
        render_items = self._render_items
        for r_item in render_items:
            real_items, rel_in_line = r_item
            if rel_in_inline:
                self._render_inline(real_items, doc)
            else:
                self._render_not_inline(real_items, doc)
    
    def _render_inline(self, real_items, doc):
        """
        同一行的对象
        """
        p = doc.add_paragraph()
        self._render_real_items(doc, p, real_items)

    def _render_not_inline(self, real_items, doc):
        """
        不是同行的对象
        """
        for real_item in real_items:
            p = doc.add_paragraph()
            self._render_real_items(p, [real_item])

    def _render_real_items(self, doc, p, real_items):
        """分发处理不同的对象"""
        for real_item in real_items:
            item_type = real_item._TYPE
            if item_type in ["image"]:
                self._render_image_item(p, real_item)
            elif item_type in ["text"]:
                self._render_text_item(p, real_item)
            elif item_type in ["table"]:
                self._render_table_item(doc, real_item)

    def _render_image_item(self, paragraph, image_item):
        """
        处理图片
        paragraph: paragraph对象
        """
        run = paragraph.add_run()
        if image_item.title:
            run.add_text(image_item.title)
        if image_item.identity:
            try:
                pic = self._get_file_stream(image_item.identity)
            except Exception as e:
                print "get_file_stream failed: %s" % e
            else:
                run.add_picture(pic)
        if image_item.text:
            run.add_text(image_item.text)

    def _render_text_item(self, paragraph, text_item):
        """
        处理文本
        """
        run = paragraph.add_run()
        if text_item.text:
            run.add_text(text_item.text)

    def _render_table_item(self, doc, table_item):
        """
        处理不同表格形式
        """
        table_type = table_item.table_type
        tables = {
            1: doc.add_table(1, 4), 
        }

        table_handlers= {
            1: self.__render_1_table,
            2: self.__render_2_table,
            3: self.__render_3_table
        }
        if tables.get(table_type) and table_handlers.get(table_type):
            table = table[table_type]
            table_handler = table_handlers[table_type]
            table_handler(table, table_item)
        else:
            print "get table or table handler wrong, type_type is %s" % table_type

    def __render_1_table(self, table, table_item):
        pass

    def __render_2_table(self, table, table_item):
        pass

    def __render_3_table(self, table, table_item):
        pass
        
    def _get_file_stream(self, file_identity):
        file_manager = self.file_manager
        file_stream = file_manager.get(file_identity)
        return file_stream.read()
