# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
## Class WattBridgeSoftware
###########################################################################

class WattBridgeSoftware ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Watt Bridge Software 1.0", pos = wx.DefaultPosition, size = wx.Size( 682,600 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		
		fgSizer1 = wx.FlexGridSizer( 0, 2, 0, 0 )
		fgSizer1.SetFlexibleDirection( wx.BOTH )
		fgSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		fgSizer2 = wx.FlexGridSizer( 0, 2, 0, 0 )
		fgSizer2.SetFlexibleDirection( wx.BOTH )
		fgSizer2.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.LineVolts = wx.StaticText( self, wx.ID_ANY, u"Line Volts", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.LineVolts.Wrap( -1 )
		fgSizer2.Add( self.LineVolts, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.LineVolts = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.LineVolts.Enable( False )
		
		fgSizer2.Add( self.LineVolts, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.LineCurrent = wx.StaticText( self, wx.ID_ANY, u"Line Current", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.LineCurrent.Wrap( -1 )
		fgSizer2.Add( self.LineCurrent, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.LineCurrent = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.LineCurrent.Enable( False )
		
		fgSizer2.Add( self.LineCurrent, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.Phase = wx.StaticText( self, wx.ID_ANY, u"Phase", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Phase.Wrap( -1 )
		fgSizer2.Add( self.Phase, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.Phase = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Phase.Enable( False )
		
		fgSizer2.Add( self.Phase, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.ActualFrequency = wx.StaticText( self, wx.ID_ANY, u"Actual Frequency", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.ActualFrequency.Wrap( -1 )
		fgSizer2.Add( self.ActualFrequency, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.ActualFrequency = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.ActualFrequency.Enable( False )
		
		fgSizer2.Add( self.ActualFrequency, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.CHHVAmpGain = wx.StaticText( self, wx.ID_ANY, u"CH HV Amp Gain", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.CHHVAmpGain.Wrap( -1 )
		fgSizer2.Add( self.CHHVAmpGain, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.CHHVAmpGain = wx.TextCtrl( self, wx.ID_ANY, u"2.497", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.CHHVAmpGain.Enable( False )
		
		fgSizer2.Add( self.CHHVAmpGain, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		fgSizer1.Add( fgSizer2, 1, wx.EXPAND, 5 )
		
		fgSizer3 = wx.FlexGridSizer( 0, 2, 0, 0 )
		fgSizer3.SetFlexibleDirection( wx.BOTH )
		fgSizer3.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.StartRow = wx.StaticText( self, wx.ID_ANY, u"Start Row", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.StartRow.Wrap( -1 )
		fgSizer3.Add( self.StartRow, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.StartRow = wx.TextCtrl( self, wx.ID_ANY, u"12", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer3.Add( self.StartRow, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.CurrentRow  = wx.StaticText( self, wx.ID_ANY, u"Current Row", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.CurrentRow .Wrap( -1 )
		fgSizer3.Add( self.CurrentRow , 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.CurrentRow = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.CurrentRow.Enable( False )
		
		fgSizer3.Add( self.CurrentRow, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.FlukeRamp = wx.StaticText( self, wx.ID_ANY, u"Fluke Ramp (s)", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.FlukeRamp.Wrap( -1 )
		fgSizer3.Add( self.FlukeRamp, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.Flukeramp = wx.TextCtrl( self, wx.ID_ANY, u"5", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer3.Add( self.Flukeramp, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.WCount = wx.StaticText( self, wx.ID_ANY, u"W Count", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.WCount.Wrap( -1 )
		fgSizer3.Add( self.WCount, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.WCount1 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.WCount1.Enable( False )
		
		fgSizer3.Add( self.WCount1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.VCount = wx.StaticText( self, wx.ID_ANY, u"V Count", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.VCount.Wrap( -1 )
		fgSizer3.Add( self.VCount, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.VCount = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.VCount.Enable( False )
		
		fgSizer3.Add( self.VCount, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		fgSizer1.Add( fgSizer3, 1, wx.EXPAND, 5 )
		
		fgSizer5 = wx.FlexGridSizer( 0, 1, 0, 0 )
		fgSizer5.SetFlexibleDirection( wx.BOTH )
		fgSizer5.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.SourceCh1 = wx.CheckBox( self, wx.ID_ANY, u"Source Ch 1", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.SourceCh1.SetValue(True) 
		fgSizer5.Add( self.SourceCh1, 0, wx.ALL, 5 )
		
		self.SourceCh2 = wx.CheckBox( self, wx.ID_ANY, u"Source Ch 2", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer5.Add( self.SourceCh2, 0, wx.ALL, 5 )
		
		self.SourceCh3 = wx.CheckBox( self, wx.ID_ANY, u"Source Ch 3", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer5.Add( self.SourceCh3, 0, wx.ALL, 5 )
		
		self.ShuntVoltsTest = wx.CheckBox( self, wx.ID_ANY, u"Shunt Volts Test", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer5.Add( self.ShuntVoltsTest, 0, wx.ALL, 5 )
		
		self.Channel1Filter = wx.CheckBox( self, wx.ID_ANY, u"Channel  1 Filter", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer5.Add( self.Channel1Filter, 0, wx.ALL, 5 )
		
		
		fgSizer1.Add( fgSizer5, 1, wx.EXPAND|wx.ALIGN_BOTTOM, 5 )
		
		fgSizer6 = wx.FlexGridSizer( 0, 2, 0, 0 )
		fgSizer6.SetFlexibleDirection( wx.BOTH )
		fgSizer6.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.Ch1TrigLevel = wx.StaticText( self, wx.ID_ANY, u"Ch 1 Trig Level", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Ch1TrigLevel.Wrap( -1 )
		fgSizer6.Add( self.Ch1TrigLevel, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		Ch1TrigLevelChoices = [ u"TTL (3 V)", u"6 V" ]
		self.Ch1TrigLevel = wx.ComboBox( self, wx.ID_ANY, u"TTL (3 V)", wx.DefaultPosition, wx.DefaultSize, Ch1TrigLevelChoices, 0 )
		self.Ch1TrigLevel.SetSelection( 0 )
		fgSizer6.Add( self.Ch1TrigLevel, 0, wx.ALL, 5 )
		
		self.Ch2TrigLevel = wx.StaticText( self, wx.ID_ANY, u"Ch 2 Trig Level", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Ch2TrigLevel.Wrap( -1 )
		fgSizer6.Add( self.Ch2TrigLevel, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		Ch2TrigLevelChoices = [ u"TTL (3 V)", u"6 V" ]
		self.Ch2TrigLevel = wx.ComboBox( self, wx.ID_ANY, u"TTL (3 V)", wx.DefaultPosition, wx.DefaultSize, Ch2TrigLevelChoices, 0 )
		self.Ch2TrigLevel.SetSelection( 0 )
		fgSizer6.Add( self.Ch2TrigLevel, 0, wx.ALL, 5 )
		
		self.SelectCounter = wx.StaticText( self, wx.ID_ANY, u"Select Counter", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.SelectCounter.Wrap( -1 )
		fgSizer6.Add( self.SelectCounter, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		SelectCounterChoices = [ u"53230A", u"53131A" ]
		self.SelectCounter = wx.ComboBox( self, wx.ID_ANY, u"53230A", wx.DefaultPosition, wx.DefaultSize, SelectCounterChoices, 0 )
		self.SelectCounter.SetSelection( 0 )
		fgSizer6.Add( self.SelectCounter, 0, wx.ALL, 5 )
		
		self.Output1 = wx.StaticText( self, wx.ID_ANY, u"52120 Output", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Output1.Wrap( -1 )
		fgSizer6.Add( self.Output1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		Output1Choices = [ u"Auto", u"High Current" ]
		self.Output1 = wx.ComboBox( self, wx.ID_ANY, u"High Current", wx.DefaultPosition, wx.DefaultSize, Output1Choices, 0 )
		self.Output1.SetSelection( 1 )
		fgSizer6.Add( self.Output1, 0, wx.ALL, 5 )
		
		self.SetDMMRangeRefVolts = wx.StaticText( self, wx.ID_ANY, u"Set DMM Range: Ref Volts", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.SetDMMRangeRefVolts.Wrap( -1 )
		fgSizer6.Add( self.SetDMMRangeRefVolts, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		SetDMMRangeRefVoltsChoices = [ u"Less than 0.7 V rms", u"Less than 7.0 V rms" ]
		self.SetDMMRangeRefVolts = wx.ComboBox( self, wx.ID_ANY, u"Less than 7.0 V rms", wx.DefaultPosition, wx.DefaultSize, SetDMMRangeRefVoltsChoices, 0 )
		self.SetDMMRangeRefVolts.SetSelection( 1 )
		fgSizer6.Add( self.SetDMMRangeRefVolts, 0, wx.ALL, 5 )
		
		self.CounterChannel = wx.StaticText( self, wx.ID_ANY, u"Counter Channel", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.CounterChannel.Wrap( -1 )
		fgSizer6.Add( self.CounterChannel, 0, wx.ALL, 5 )
		
		CounterChannelChoices = [ u"Both Channels", u"Channel 1 only" ]
		self.CounterChannel = wx.ComboBox( self, wx.ID_ANY, u"Channel 1 only", wx.DefaultPosition, wx.DefaultSize, CounterChannelChoices, 0 )
		self.CounterChannel.SetSelection( 1 )
		fgSizer6.Add( self.CounterChannel, 0, wx.ALL, 5 )
		
		self.OutputAutoHigh = wx.StaticText( self, wx.ID_ANY, u"52120 Output", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.OutputAutoHigh.Wrap( -1 )
		fgSizer6.Add( self.OutputAutoHigh, 0, wx.ALL, 5 )
		
		OutputAutoHighChoices = [ u"AUTO", u"HIGH" ]
		self.OutputAutoHigh = wx.ComboBox( self, wx.ID_ANY, u"AUTO", wx.DefaultPosition, wx.DefaultSize, OutputAutoHighChoices, 0 )
		self.OutputAutoHigh.SetSelection( 0 )
		fgSizer6.Add( self.OutputAutoHigh, 0, wx.ALL, 5 )
		
		
		fgSizer1.Add( fgSizer6, 1, wx.EXPAND, 5 )
		
		fgSizer7 = wx.FlexGridSizer( 0, 1, 0, 0 )
		fgSizer7.SetFlexibleDirection( wx.BOTH )
		fgSizer7.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.StartNewSequence = wx.Button( self, wx.ID_ANY, u"Start New Sequence (from \"Start Row\")", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer7.Add( self.StartNewSequence, 0, wx.ALL, 5 )
		
		self.ContinueSequence = wx.Button( self, wx.ID_ANY, u"Continue Sequence (from \"Current Row\")", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer7.Add( self.ContinueSequence, 0, wx.ALL, 5 )
		
		self.SaveData = wx.Button( self, wx.ID_ANY, u"Save Data", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer7.Add( self.SaveData, 0, wx.ALL, 5 )
		
		self.MakeSafe = wx.Button( self, wx.ID_ANY, u"Make Safe", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer7.Add( self.MakeSafe, 0, wx.ALL, 5 )
		
		self.CheckConnections = wx.Button( self, wx.ID_ANY, u"Check Connections", wx.DefaultPosition, wx.DefaultSize, 0 )
		fgSizer7.Add( self.CheckConnections, 0, wx.ALL, 5 )
		
		
		fgSizer1.Add( fgSizer7, 1, wx.EXPAND, 5 )
		
		fgSizer10 = wx.FlexGridSizer( 0, 1, 0, 0 )
		fgSizer10.SetFlexibleDirection( wx.BOTH )
		fgSizer10.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.WattBridgeEventsLog = wx.StaticText( self, wx.ID_ANY, u"Watt Bridge Events Log", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.WattBridgeEventsLog.Wrap( -1 )
		fgSizer10.Add( self.WattBridgeEventsLog, 0, wx.ALL, 5 )
		
		self.WattBridgeEventsLog = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 400,170 ), wx.TE_MULTILINE )
		fgSizer10.Add( self.WattBridgeEventsLog, 0, wx.ALL, 5 )
		
		
		fgSizer1.Add( fgSizer10, 1, wx.EXPAND, 5 )
		
		
		self.SetSizer( fgSizer1 )
		self.Layout()
		self.m_menubar1 = wx.MenuBar( 0 )
		self.File = wx.Menu()
		self.About = wx.MenuItem( self.File, wx.ID_ANY, u"About", wx.EmptyString, wx.ITEM_NORMAL )
		self.File.AppendItem( self.About )
		
		self.m_menubar1.Append( self.File, u"File" ) 
		
		self.SetMenuBar( self.m_menubar1 )
		
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.Bind( wx.EVT_CLOSE, self.WattBridgeSoftwareOnClose )
		self.StartNewSequence.Bind( wx.EVT_BUTTON, self.StartNewSequenceOnButtonClick )
		self.ContinueSequence.Bind( wx.EVT_BUTTON, self.ContinueSequenceOnButtonClick )
		self.SaveData.Bind( wx.EVT_BUTTON, self.SaveDataOnButtonClick )
		self.MakeSafe.Bind( wx.EVT_BUTTON, self.MakeSafeOnButtonClick )
		self.CheckConnections.Bind( wx.EVT_BUTTON, self.CheckConnectionsOnButtonClick )
		self.Bind( wx.EVT_MENU, self.AboutOnMenuSelection, id = self.About.GetId() )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def WattBridgeSoftwareOnClose( self, event ):
		event.Skip()
	
	def StartNewSequenceOnButtonClick( self, event ):
		event.Skip()
	
	def ContinueSequenceOnButtonClick( self, event ):
		event.Skip()
	
	def SaveDataOnButtonClick( self, event ):
		event.Skip()
	
	def MakeSafeOnButtonClick( self, event ):
		event.Skip()
	
	def CheckConnectionsOnButtonClick( self, event ):
		event.Skip()
	
	def AboutOnMenuSelection( self, event ):
		event.Skip()
	

###########################################################################
## Class About
###########################################################################

class About ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"About", pos = wx.DefaultPosition, size = wx.Size( 304,223 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		
		fgSizer8 = wx.FlexGridSizer( 0, 1, 0, 0 )
		fgSizer8.SetFlexibleDirection( wx.BOTH )
		fgSizer8.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.Program = wx.StaticText( self, wx.ID_ANY, u"Program: Watt Bridge Software ", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Program.Wrap( -1 )
		fgSizer8.Add( self.Program, 0, wx.ALL, 5 )
		
		self.Version = wx.StaticText( self, wx.ID_ANY, u"Version: 1.0", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Version.Wrap( -1 )
		fgSizer8.Add( self.Version, 0, wx.ALL, 5 )
		
		self.Author = wx.StaticText( self, wx.ID_ANY, u"Author: Ikram Singh", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Author.Wrap( -1 )
		fgSizer8.Add( self.Author, 0, wx.ALL, 5 )
		
		self.Organization = wx.StaticText( self, wx.ID_ANY, u"MSL @ Callaghan Innovation ", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.Organization.Wrap( -1 )
		fgSizer8.Add( self.Organization, 0, wx.ALL, 5 )
		
		
		self.SetSizer( fgSizer8 )
		self.Layout()
		
		self.Centre( wx.BOTH )
	
	def __del__( self ):
		pass
	

