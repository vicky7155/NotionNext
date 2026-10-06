# -*- coding: utf-8 -*-
"""GEA 官网简介文案（正式版）——库函数版"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
HEI, FANG, SONG = u'黑体', u'仿宋', u'宋体'

def _set_font(run, name, size, bold=False):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run.font.size = Pt(size); run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)

def build(out_root):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54); sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(2.8); sec.right_margin = Cm(2.6)

    def center(text, font, size, bold=False, before=0, after=6):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
        r = p.add_run(text); _set_font(r, font, size, bold)

    def h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(6)
        r = p.add_run(text); _set_font(r, HEI, 14, True)

    def para(text, size=12, font=FANG, bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.6; p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.first_line_indent = Cm(0.9)
        r = p.add_run(text); _set_font(r, font, size, bold)

    def li(text, size=12):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5; p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.9)
        r = p.add_run(text); _set_font(r, FANG, size)

    center(u'全球创业者联盟（GEA）官网简介文案', HEI, 20, True, 6, 4)
    center(u'（正式版　V1.0　·　用途：品牌官方网站）', HEI, 11, False, 0, 2)
    center(u'发布单位：全球创业者联盟总部', FANG, 10.5, False, 0, 14)

    h2(u'【首屏主视觉区】')
    center(u'全球创业者联盟', HEI, 22, True, 8, 2)
    center(u'GLOBAL ENTREPRENEURS ALLIANCE', SONG, 11, False, 0, 8)
    center(u'让每一位创业者，都能找到一个可以信任的人', HEI, 15, True, 0, 8)
    center(u'2019年成立于美国旧金山　|　面向全球206个国家与地区　|　规则先于规模', FANG, 11, False, 0, 6)
    center(u'〔按钮：申请入会〕　〔按钮：登记项目／资源〕', SONG, 10.5, False, 0, 4)

    h2(u'【第一板块　GEA 是什么】')
    para(u'全球创业者联盟（Global Entrepreneurs Alliance，GEA）是一个国际创业资源对接组织。', bold=True)
    para(u'我们只做一件事：把有项目的人和有资源的人，在统一的规则与全程的保障之下，对接起来，促成真实、安全的跨境合作。')
    para(u'GEA 于2019年在美国旧金山成立，以官方徽章为信，以「规则先于规模」为立盟之本，面向全球206个国家与地区逐步设立城市分会，连接创业者、投资人、项目方、资源方、服务商与分会，共建六方共赢的跨境创业生态。')

    h2(u'【第二板块　GEA 做什么——三大核心业务】')
    para(u'一、资源撮合', bold=True)
    para(u'你有项目，我们帮你对接资金、渠道、技术、供应链；你有资源，我们帮你找到经过审查的可靠项目。跨境需求先登记、分会先初审、总部再复核，按项目评分分级精准匹配。')
    para(u'二、交易保障', bold=True)
    para(u'撮合不是拉个群、做个介绍。每一单合作都按26份全球统一执行文件推进：资格审查、立项评估、合同签署、里程碑验收，环环有文件、步步可追溯；履约保障金由第三方机构专户托管，节点验收通过才按约释放。')
    para(u'三、信用生态', bold=True)
    para(u'会员实名审核入会，每笔合作双向评价并计入信用档案；八条合规红线，触碰即除名、永不复籍。在GEA，"可信任"不是口号，是看得见、查得到的记录。')

    h2(u'【第三板块　怎么运作——四步闭环】')
    li(u'第一步　登记需求：会员实名入会，项目方与资源方分别通过资格审查，需求统一登记；')
    li(u'第二步　评估匹配：项目评分分级，分会初审、总部复核，智能匹配候选资源，双方确认意向；')
    li(u'第三步　履约管护：三方签署主合同，保障金第三方托管，按里程碑节点交付验收，风险实时预警；')
    li(u'第四步　结算复盘：收益三方签章结算，服务费只在成交之后收取，双向评价计入信用档案。')

    h2(u'【第四板块　为什么选择 GEA——五个承诺】')
    li(u'■ 规则先于规模：19项管理制度、26份全球统一执行文件，先有规则，再谈生意；')
    li(u'■ 钱是安全的：履约保障金第三方专户托管，包括联盟在内，任何人无权单方动用；')
    li(u'■ 人是审过的：实名入会、双重资格审查、信用档案全程记录；')
    li(u'■ 事是可追溯的：从登记到结算，环环有文件、事事有记录；')
    li(u'■ 不成交，不收费：服务费只在成交之后收取。')

    h2(u'【第五板块　结尾号召区】')
    para(u'GEA 不做只办活动的圈子，而是建成一套能让跨境生意真正发生的制度与体系。')
    para(u'无论你是正在找资源的创业者、寻找好项目的投资人，还是希望在你的城市设立分会的同行者——')
    center(u'让每一位创业者，都能找到一个可以信任的人。', HEI, 13, True, 6, 6)
    center(u'〔按钮：申请入会〕　〔按钮：申请设立城市分会〕　〔按钮：联系我们〕', SONG, 10.5, False, 0, 4)

    h2(u'【使用说明（给网站编辑，上线时删除本段）】')
    li(u'1．【首屏】至【第五板块】按顺序对应官网首页从上到下的五个区块；', 10.5)
    li(u'2．〔按钮：××〕为行动按钮占位，链接到对应功能页；', 10.5)
    li(u'3．正文语句可直接使用，不建议删改"三大核心业务"与"五个承诺"的结构；', 10.5)
    li(u'4．配图建议：首屏用全球地图／城市天际线，板块二用图标式三栏，板块三用四步流程图；', 10.5)
    li(u'5．随文件包附带 preview.html，为本文案的网页排版示例，可交给前端直接参考实现。', 10.5)

    out = os.path.join(out_root, u'全球创业者联盟GEA官网简介文案（正式版V1.0）.docx')
    doc.save(out)
    print('web intro saved')
