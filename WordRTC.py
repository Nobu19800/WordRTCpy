#!/bin/env python
# -*- encoding: cp932 -*-

##
#   @file WordRTC.py
#   @brief WordControl Component

import win32com
import pythoncom
import pdb
from win32com.client import *
import pprint
import datetime
import msvcrt




import optparse
import sys,os,platform
import re
import time
import random
import commands
import math



import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming

from WriterControl import *


wordcontrol_spec = ["implementation_id", "WordControl",
                  "type_name",         "WordControl",
                  "description",       "Word Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.file_path", "NewFile",
                  "conf.default.fontsize", "16",
                  #"conf.default.fontname", "ＭＳ 明朝",
                  "conf.default.char_Red", "0",
                  "conf.default.char_Blue", "0",
                  "conf.default.char_Green", "0",
                  "conf.default.italic", "0",
                  "conf.default.bold", "0",
                  "conf.default.underline", "0",
                  "conf.default.shadow", "0",
                  "conf.default.strikeout", "0",
                  "conf.default.contoured", "0",
                  "conf.default.emphasis", "0",
                  "conf.default.back_Red", "255",
                  "conf.default.back_Blue", "255",
                  "conf.default.back_Green", "255",
                  "conf.default.code", "utf-8",
                  "conf.__widget__.file_path", "text",
                  "conf.__widget__.fontsize", "spin",
                  #"conf.__widget__.fontname", "radio",
                  "conf.__widget__.char_Red", "spin",
                  "conf.__widget__.char_Blue", "spin",
                  "conf.__widget__.char_Green", "spin",
                  "conf.__widget__.italic", "radio",
                  "conf.__widget__.bold", "radio",
                  "conf.__widget__.underline", "radio",
                  "conf.__widget__.shadow", "radio",
                  "conf.__widget__.strikeout", "radio",
                  "conf.__widget__.contoured", "radio",
                  "conf.__widget__.emphasis", "radio",
                  "conf.__widget__.back_Red", "spin",
                  "conf.__widget__.back_Blue", "spin",
                  "conf.__widget__.back_Green", "spin",
                  "conf.__widget__.code", "radio",
                  "conf.__constraints__.fontsize", "1<=x<=72",
                  #"conf.__constraints__.fontname", "(MS UI Gothic,MS ゴシック,MS Pゴシック,MS 明朝,MS P明朝,HG ゴシック E,HGP ゴシック E,HGS ゴシック E,HG ゴシック M,HGP ゴシック M,HGS ゴシック M,HG 正楷書体-PRO,HG 丸ゴシック M-PRO,HG 教科書体,HGP 教科書体,HGS 教科書体,HG 行書体,HGP 行書体,HGS 行書体,HG 創英プレゼンス EB,HGP 創英プレゼンス EB,HGS 創英プレゼンス EB,HG 創英角ゴシック UB,HGP 創英角ゴシック UB,HGS 創英角ゴシック UB,HG 創英角ポップ体,HGP 創英角ポップ体,HGS 創英角ポップ体,HG 明朝 B,HGP 明朝 B,HGS 明朝 B,HG 明朝 E,HGP 明朝 E,HGS 明朝 E,メイリオ)",
                  "conf.__constraints__.char_Red", "0<=x<=255",
                  "conf.__constraints__.char_Blue", "0<=x<=255",
                  "conf.__constraints__.char_Green", "0<=x<=255",
                  "conf.__constraints__.italic", "(0,1)",
                  "conf.__constraints__.bold", "(0,1)",
                  "conf.__constraints__.underline", "(0,1)",
                  "conf.__constraints__.shadow", "(0,1)",
                  "conf.__constraints__.strikeout", "(0,1)",
                  "conf.__constraints__.contoured", "(0,1)",
                  "conf.__constraints__.emphasis", "(0,1)",
                  "conf.__constraints__.back_Red", "0<=x<=255",
                  "conf.__constraints__.back_Blue", "0<=x<=255",
                  "conf.__constraints__.back_Green", "0<=x<=255",
                  "conf.__constraints__.code", "(utf-8,euc_jp,shift_jis)",
                  ""]




##
# @class WordConfigUpdateParam
# @brief コンフィギュレーションパラメータが更新されたときのコールバック
#

class WordConfigUpdateParam(OpenRTM_aist.ConfigurationSetListener):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param e_rtc OOoCalcRTC
    #
   def __init__(self,e_rtc):
        self.m_rtc =  e_rtc

    ##
    # @brief 
    # @param self 
    # @param config_param_name 
    #
   def __call__(self, config_param_name):
        self.m_rtc.configUpdate()

   


##
# @class WordControl
# @brief Wordを操作するためのRTCのクラス
#

class WordControl(WriterControl):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    WriterControl.__init__(self, manager)
    
    prop = OpenRTM_aist.Manager.instance().getConfig()
    fn = self.getProperty(prop, "word.filename", "")
    self.m_word = WordObject()
    if fn != "":
      str1 = [fn]
      OpenRTM_aist.replaceString(str1,"/","\\")
      fn = os.path.abspath(str1[0])
    self.m_word.Open(fn)

    self.conf_filename = ["NewFile"]
    
    
    
    
    
    return

  ##
  # @brief rtc.confの設定を取得する関数
  #
  def getProperty(self, prop, key, value):
        
        if  prop.findNode(key) != None:
            #print value
            value = prop.getProperty(key)
        return value
    
  ##
  # @brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
  # @param self 
  #
  def configUpdate(self):
      return
      """self._configsets.update("default","file_path")
      str1 = [self.conf_filename[0]]
      OpenRTM_aist.replaceString(str1,"/","\\")
      sfn = str1[0]
      tfn = os.path.abspath(sfn)
      self.m_word.initCom()
      if sfn == "NewFile":
        self.m_word.Open("")
      else:
        print sfn,tfn
        self.m_word.Open(tfn)"""

      #self.m_word.closeCom()
        
      
  
  ##
  # @brief 文字書き込みの関数
  # @param self 
  # @param m_str 書き込む文字列
  #

  def setWord(self, m_str):
      self.m_word.setFontColor(self.char_Red, self.char_Green, self.char_Blue)
      self.m_word.setBackColor(self.back_Red, self.back_Green, self.back_Blue)
      self.m_word.setFontSize(self.fontSize)
      self.m_word.setFontName(self.fontName)
      self.m_word.bold = self.bold
      #print SetCoding(m_str, self.conf_Code[0])
      self.m_word.setWord(SetCoding(m_str, self.conf_Code[0]))


  ##
  # @brief カーソル位置の文字取得の関数
  # @param self
  # @return カーソル位置の文字列
  #

  def getWord(self):
      return self.m_word.getSelWord()


  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  def onInitialize(self):
    WriterControl.onInitialize(self)

    self.bindParameter("file_path", self.conf_filename, "NewFile")

    self.addConfigurationSetListener(OpenRTM_aist.ConfigurationSetListenerType.ON_SET_CONFIG_SET, WordConfigUpdateParam(self))
    
    return RTC.RTC_OK
   
  ##
  # @brief 活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onActivated(self, ec_id):
    WriterControl.onActivated(self, ec_id)
    
    

    #self.file = open('text3.txt', 'w')

    self.m_word.initCom()
    
    return RTC.RTC_OK


  ##
  # @brief 不活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onDeactivated(self, ec_id):
    WriterControl.onDeactivated(self, ec_id)
    #self.m_word.closeCom()
    #self.file.close()
    return RTC.RTC_OK


  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    WriterControl.onExecute(self, ec_id)
        

    return RTC.RTC_OK

  
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def on_shutdown(self, ec_id):
      WriterControl.on_shutdown(self, ec_id)
      return RTC.RTC_OK


  
  
  
      

  ##
  # @brief 文字数移動する関数
  # @param self 
  # @param diff 移動する文字数
  #
  def moveCharacter(self, diff):
      self.m_word.moveSelection(WordObject.wdCharacter,diff)
          
  ##
  # @brief 単語数移動する関数
  # @param self 
  # @param diff 移動する単語数
  #
  def moveWord(self, diff):
      self.m_word.moveSelection(WordObject.wdWord,diff)

  ##
  # @brief 行数移動する関数
  # @param self 
  # @param diff 移動する行数
  #
  def moveLine(self, diff):
      self.m_word.moveSelection(WordObject.wdLine,diff)

  ##
  # @brief 段落数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveParagraph(self, diff):
      self.m_word.moveSelection(WordObject.wdParagraph,diff)

  ##
  # @brief ウインドウ数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveWindow(self, diff):
      self.m_word.moveSelection(WordObject.wdWindow,diff)

  ##
  # @brief スクリーン数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveScreen(self, diff):
      self.m_word.moveSelection(WordObject.wdScreen,diff)



##
# @class WordObject
# @brief Wordを操作するクラス
#
class WordObject:
    wdCharacter = 1
    wdWord = 2
    wdSentence = 3
    wdParagraph = 4
    wdLine = 5
    wdStory = 6
    wdScreen = 7
    wdSection = 8
    wdColumn = 9
    wdRow = 10
    wdWindow = 11
    wdTable = 15

    wdMove = 0
    wdExtend = 1

    ##
    # @brief コンストラクタ
    # @param self 
    #
    def __init__(self):
        self.dcApplication = None
        self.dcDocuments = None
        self.dcDocument = None
        self.fontSize = 20
        self.fontName = "ＭＳ 明朝"
        self.filename = " "
        self.char_Red = 0
        self.char_Green = 0
        self.char_Blue = 0
        self.movementType = False
        self.italic = False
        self.bold = False

        self.underline = False
        self.shadow = False
        self.strikeout = False
        self.contoured = False
        self.emphasis = False

        self.back_Red = 255
        self.back_Green = 255
        self.back_Blue = 255

        self.thread_dcApplication = None
        self.thread_dcDocuments = None
        self.thread_dcDocument = None

        self.t_dcApplication = None
        self.t_dcDocuments = None
        self.t_dcDocument = None

    ##
    # @brief 文字の色を変更する関数
    # @param self
    # @param r 文字の色(R)
    # @param g 文字の色(G)
    # @param b 文字の色(B)
    #
    def setFontColor(self, r, g, b):
        self.char_Red = r
        self.char_Green = g
        self.char_Blue = b

    ##
    # @brief 背景色を変更する関数
    # @param self
    # @param r 文字の色(R)
    # @param g 文字の色(G)
    # @param b 文字の色(B)
    #
    def setBackColor(self, r, g, b):
        self.back_Red = r
        self.back_Green = g
        self.back_Blue = b

    ##
    # @brief フォントのサイズを変更する関数
    # @param self
    # @param fs フォントサイズ
    #
    def setFontSize(self, fs):
        self.fontSize = fs

    ##
    # @brief フォントの種類を変更する関数
    # @param self
    # @param fn フォントの種類
    #
    def setFontName(self, fn):
        self.fontName = fn

    ##
    # @brief 文字列を書き込む関数
    # @param self
    # @param st 書き込む文字列
    #
    def setWord(self, st):
        
        
        currentSelection = self.dcApplication.Selection
        

        slc = currentSelection.End

        
        tr = self.dcDocument.Range(slc,slc)
        

        tr.Text = st

        tr.Font.Size = self.fontSize
        tr.Font.Name = self.fontName
        tr.Font.Bold = self.bold

        currentSelection.MoveRight(WordObject.wdCharacter, len(st) , WordObject.wdMove)

    ##
    # @brief 選択位置の文字列を取得する関数
    # @param self
    # @return 取得した文字列
    #
    def getSelWord(self):

        currentSelection = self.dcApplication.Selection
        
        return currentSelection.Text.encode("utf-8")

    ##
    # @brief 選択位置を変更する関数
    # @param self
    # @param m_type 移動のタイプ
    # @param leng 移動量
    #
    def moveSelection(self, m_type, leng):
        currentSelection = self.dcApplication.Selection
        if not self.MovementType:
            currentSelection.MoveRight(m_type, leng, WordObject.wdMove)
        else:
            currentSelection.MoveRight(m_type, leng, WordObject.wdExtend)

    ##
    # @brief 
    # @param self
    #
    def preInitCom(self):
        self.thread_dcApplication = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_dcApplication)
        self.thread_dcDocuments = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_dcDocuments)
        self.thread_dcDocument = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_dcDocument)

    ##
    # @brief 
    # @param self
    #
    def initCom(self):
        #if self.dcApplication == None:
           pythoncom.CoInitialize()
           self.dcApplication = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_dcApplication, pythoncom.IID_IDispatch))
           self.dcDocuments = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_dcDocuments, pythoncom.IID_IDispatch))
           self.dcDocument = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_dcDocument, pythoncom.IID_IDispatch))

    ##
    # @brief 
    # @param self
    #
    def closeCom(self):
        pythoncom.CoUninitialize()
        self.dcApplication = None
        self.dcDocuments = None
        self.dcDocument = None
        

    ##
    # @brief Wordファイルを開く関数
    # @param self
    # @param fn ファイルパス
    #
    def Open(self, fn):
        if self.filename == fn:
            return
        self.filename = fn

        
        
        try:
            
            
            t_dcApplication = win32com.client.Dispatch("Word.Application")
               
            
              
            
            
            
            t_dcApplication.Visible = True
            try:
                t_dcDocuments = t_dcApplication.Documents
                

                try:
                    t_dcDocument = None
                    if self.filename == "":
                        t_dcDocument = t_dcDocuments.Add()
                    else:
                        t_dcDocument = t_dcDocuments.Open(self.filename)

                    
                    if self.t_dcApplication == None:
                       self.t_dcApplication = t_dcApplication
                       self.t_dcDocuments = t_dcDocuments
                       self.t_dcDocument = t_dcDocument
                       self.preInitCom()
                    
                    
                    

                    

                    
                except:
                    return
            except:
                return
        except:
            return


##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=wordcontrol_spec)
    manager.registerFactory(profile,
                            WordControl,
                            OpenRTM_aist.Delete)
    comp = manager.createComponent("WordControl")

def main():
    
    
    
    mgr = OpenRTM_aist.Manager.init(sys.argv)
    mgr.setModuleInitProc(MyModuleInit)
    mgr.activateManager()
    mgr.runManager()

    
if __name__ == "__main__":
    main()
