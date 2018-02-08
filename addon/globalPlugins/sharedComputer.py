# -*- coding: UTF-8 -*-
# sharedComputer: Plugin to using shared computers
# https://github.com/nvaccess/nvda/issues/4093
# https://github.com/nvaccess/nvda/pull/7506/files
# Copyright (C) 2018 Robert Hänggi, Noelia Ruiz Martínez
# Released under GPL 2

import globalPluginHandler
import addonHandler
import gui, ui, wx, time, tones, winUser, config
from gui import guiHelper
from gui import nvdaControls
from gui.settingsDialogs import SettingsDialog
from keyboardHandler import KeyboardInputGesture
from globalCommands import SCRCAT_CONFIG
from comtypes import HRESULT, GUID, IUnknown, COMMETHOD, POINTER, CoCreateInstance, cast, c_float
from ctypes.wintypes import BOOL, DWORD, UINT 
from logHandler import log
from api import processPendingEvents

addonHandler.initTranslation()

### Constants
ADDON_NAME = str(addonHandler.getCodeAddon().name)

numLockActivationDefault = "0" if config.conf['keyboard']['keyboardLayout'] == "desktop" else "2"

confspec = {
	"numLockActivation": "integer(default="+numLockActivationDefault+")",
	"changeVolumeLevel": "integer(default=0)",
	"volumeLevel": "integer(default=50)"
}
config.conf.spec[ADDON_NAME] = confspec

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def handleConfigProfileSwitch(self):
		numLockActivation = config.conf[ADDON_NAME]["numLockActivation"]
		if numLockActivation < 2 and winUser.getKeyState(winUser.VK_NUMLOCK) != numLockActivation:
			KeyboardInputGesture.fromName("numLock").send()

	def changeVolumeLevel(self, targetLevel, mode):
		if self.speakers:
			# should actually work on first attempt
			# but level 0 is a pain
			for attempt in range(2):
				processPendingEvents()
				level = int(self.speakers.GetMasterVolumeLevelScalar()*100)
				log.info("Level speakers at Startup: {oldLevel} Percent".format(oldLevel=level))
				if level < targetLevel or (mode == 1 and self.level > targetLevel):
					self.speakers.SetMasterVolumeLevelScalar(targetLevel/100.0, None)
				muteState = self.speakers.GetMute()
				log.info("speakers at Startup: {curMuteState}".format(curMuteState=("Unmuted", "Muted")[muteState]))
				if muteState:
					self.speakers.SetMute(0, None)
				time.sleep(0.05)
				log.info("speakers after correction: {fixedLevel} Percent, {curMuteState}".format(
					fixedLevel=int(self.speakers.GetMasterVolumeLevelScalar()*100),
					curMuteState=("Unmuted","Muted")[muteState]))
		else:
			# As a fall-back, change the volume "manually"
			# ensures only a minimum
			volDown = KeyboardInputGesture.fromName("VolumeDown")
			volUp = KeyboardInputGesture.fromName("VolumeUp")
			# one keystroke = 2 % here, is that universal?
			repeats = targetLevel//2
			# We are only interested in the side effect, hence the dummy underscore variable
			_ = {key.send() for key in iter((volDown,)*repeats + (volUp,)*repeats)}
			time.sleep(float(level)/62)

	def __init__(self):
		self.speakers=getVolumeObject()
		super(globalPluginHandler.GlobalPlugin, self).__init__()
		volLevel = config.conf[ADDON_NAME]["volumeLevel"]
		volMode = config.conf[ADDON_NAME]["changeVolumeLevel"]
		if volMode < 2:
			wx.CallAfter(self.changeVolumeLevel, volLevel, volMode)
		self.numLockState = winUser.getKeyState(winUser.VK_NUMLOCK)
		self.handleConfigProfileSwitch()
		try:
			config.configProfileSwitched.register(self.handleConfigProfileSwitch)
		except AttributeError:
			pass

		# Gui
		self.prefsMenu = gui.mainFrame.sysTrayIcon.preferencesMenu
		self.settingsItem = self.prefsMenu.Append(wx.ID_ANY,
			# Translators: Name of an option in the menu.
			_("&Shared Computer settings..."))
		gui.mainFrame.sysTrayIcon.Bind(wx.EVT_MENU, self.onSettings, self.settingsItem)

	def terminate(self):
		if winUser.getKeyState(winUser.VK_NUMLOCK) != self.numLockState:
			KeyboardInputGesture.fromName("numLock").send()
		try:
			config.configProfileSwitched.unregister(self.handleConfigProfileSwitch)
		except AttributeError: # This is for backward compatibility
			pass
		try:
			self.prefsMenu.RemoveItem(self.settingsItem)
		except: # Compatible with Python 2 and 3
			pass

	def onSettings(self, evt):
		gui.mainFrame._popupSettingsDialog(AddonSettingsDialog)

	def script_settings(self, gesture):
		wx.CallAfter(self.onSettings, None)
	script_settings.category = SCRCAT_CONFIG
	# Translators: message presented in input mode.
	script_settings.__doc__ = _("Shows the Shared Computer settings dialog.")

class AddonSettingsDialog(SettingsDialog):

	# Translators: Title of a dialog.
	title = _("Shared Computer settings")

	def makeSettings(self, settingsSizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
		# Translators: label of a dialog.
		activateLabel = _("&Activate NumLock:")
		self.activateChoices = (
			# Translators: Choice of a dialog.
			_("Off"),
			# Translators: Choice of a dialog.
			_("On"),
			# Translators: Choice of a dialog.
			_("Never change"))
		self.activateList = sHelper.addLabeledControl(activateLabel, wx.Choice, choices=self.activateChoices)
		self.activateList.Selection = config.conf[ADDON_NAME]["numLockActivation"]
		# Translators: label of a dialog.
		volumeLabel = _("System &Volume at Start:")
		self.volumeChoices = (
			# Translators: Choice of a dialog.
			_("Ensure a minimum of"),
			# Translators: Choice of a dialog.
			_("Set exactly to"),
			# Translators: Choice of a dialog.
			_("Never change"))
		self.volumeList = sHelper.addLabeledControl(volumeLabel, wx.Choice, choices=self.volumeChoices)
		self.volumeList.Selection = config.conf[ADDON_NAME]["changeVolumeLevel"]
		self.volumeList.Bind(wx.EVT_CHOICE, self.onChoice) 
		# Translators: Label of a dialog.
		initialVolumeLabel = _("Initial value for volume level:")
		initialVolumeChoices = (
			# Translators: Choice of a dialog.
			_("Set in the add-on &configuration"),
			# Translators: Choice of a dialog.
			_("Current &system volume level")
		)
		self.volumeRadioBox=sHelper.addItem(wx.RadioBox(self, label=initialVolumeLabel, choices=initialVolumeChoices))
		self.volumeRadioBox.Bind(wx.EVT_RADIOBOX, self.onVolumeRadioBox)
		# Translators: Label of a dialog.
		self.volumeLevel = sHelper.addLabeledControl(_("Volume &Level:"), 
			nvdaControls.SelectOnFocusSpinCtrl, 
			min = 20 if self.volumeList.Selection==1 else 0, 
			initial=config.conf[ADDON_NAME]["volumeLevel"])

	def onChoice(self, evt):
		val=evt.GetSelection()
		if val == 0:
			self.volumeLevel.SetRange(0, 100)
			self.volumeLevel.Enabled=True
		elif val==1:
			self.volumeLevel.SetRange(20, 100)
			self.volumeLevel.Enabled=True
		else:
			self.volumeLevel.Enabled=False

	def onVolumeRadioBox(self, evt):
		if self.volumeRadioBox.Selection == 1:
			try:
				speakers = getVolumeObject()
				self.volumeLevel.Value = int(speakers.GetMasterVolumeLevelScalar()*100)
			except:
				self.volumeLevel.Value = 1
				#self.volumeLevel.Value = config.conf[ADDON_NAME]["volumeLevel"]
		else:
			self.volumeLevel.Value = config.conf[ADDON_NAME]["volumeLevel"]

	def postInit(self):
		self.activateList.SetFocus()

	def onOk(self, evt):
		super(AddonSettingsDialog, self).onOk(evt)
		config.conf[ADDON_NAME]["numLockActivation"] = self.activateList.Selection
		# We cannot write only to the normal configuration, since this produces a key error
		config.conf[ADDON_NAME]["changeVolumeLevel"] = self.volumeList.Selection
		config.conf[ADDON_NAME]["volumeLevel"] = self.volumeLevel.Value

### Audio Stuff

def getVolumeObject():
	CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')
	deviceEnumerator = CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, 1)
	volObj = cast(
		deviceEnumerator.GetDefaultAudioEndpoint(0, 1).Activate(IAudioEndpointVolume._iid_, 7, None),
		POINTER(IAudioEndpointVolume))
	return volObj

# for a ffull-fletched Audio wrapper
# visit https://github.com/AndreMiras/pycaw
class IAudioEndpointVolume(IUnknown):
	_iid_ = GUID('{5CDF2C82-841E-4546-9722-0CF74078229A}')
	_methods_ = (
		COMMETHOD([], HRESULT, 'NotImpl1'),
		COMMETHOD([], HRESULT, 'NotImpl2'),
		COMMETHOD([], HRESULT, 'GetChannelCount', (['out'], POINTER(UINT), 'pnChannelCount')),
		COMMETHOD([], HRESULT, 'SetMasterVolumeLevel',
			(['in'], c_float, 'fLevelDB'), (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'SetMasterVolumeLevelScalar',
			(['in'], c_float, 'fLevel'), (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'GetMasterVolumeLevel', (['out'], POINTER(c_float), 'pfLevelDB')),
		COMMETHOD([], HRESULT, 'GetMasterVolumeLevelScalar', (['out'], POINTER(c_float), 'pfLevelDB')),
		COMMETHOD([], HRESULT, 'SetChannelVolumeLevel',
			(['in'], UINT, 'nChannel'), (['in'], c_float, 'fLevelDB'), (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'SetChannelVolumeLevelScalar',
			(['in'], DWORD, 'nChannel'), (['in'], c_float, 'fLevelDB'), (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'GetChannelVolumeLevel',
			(['in'], UINT, 'nChannel'),
			(['out'], POINTER(c_float), 'pfLevelDB')),
		COMMETHOD([], HRESULT, 'GetChannelVolumeLevelScalar',
			(['in'], DWORD, 'nChannel'),
			(['out'], POINTER(c_float), 'pfLevelDB')),
		COMMETHOD([], HRESULT, 'SetMute', (['in'], BOOL, 'bMute'), (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'GetMute', (['out'], POINTER(BOOL), 'pbMute')),
		COMMETHOD([], HRESULT, 'GetVolumeStepInfo',
			(['out'], POINTER(DWORD), 'pnStep'),
			(['out'], POINTER(DWORD), 'pnStepCount')),
		COMMETHOD([], HRESULT, 'VolumeStepUp', (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'VolumeStepDown', (['in'], POINTER(GUID), 'pguidEventContext')),
		COMMETHOD([], HRESULT, 'QueryHardwareSupport', (['out'], POINTER(DWORD), 'pdwHardwareSupportMask')),
		COMMETHOD([], HRESULT, 'GetVolumeRange',
			(['out'], POINTER(c_float), 'pfMin'),
			(['out'], POINTER(c_float), 'pfMax'),
			(['out'], POINTER(c_float), 'pfIncr')))

class IMMDevice(IUnknown):
	_iid_ = GUID('{D666063F-1587-4E43-81F1-B948E807363F}')
	_methods_ = (
		COMMETHOD([], HRESULT, 'Activate',
			(['in'], POINTER(GUID), 'iid'),
			(['in'], DWORD, 'dwClsCtx'),
			(['in'], POINTER(DWORD), 'pActivationParams'),
			(['out'], POINTER(POINTER(IUnknown)), 'ppInterface')),)

class IMMDeviceCollection(IUnknown):
	_iid_ = GUID('{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}')
	_methods_ = (
		COMMETHOD([], HRESULT, 'GetCount',
			(['out'], POINTER(UINT), 'pcDevices')),
		COMMETHOD([], HRESULT, 'Item',
			(['in'], UINT, 'nDevice'),
			(['out'], POINTER(POINTER(IMMDevice)), 'ppDevice')))

class IMMDeviceEnumerator(IUnknown):
	_iid_ = GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
	_methods_ = (
		COMMETHOD([], HRESULT, 'EnumAudioEndpoints',
			(['in'], DWORD, 'dataFlow'),
			(['in'], DWORD, 'dwStateMask'),
			(['out'], POINTER(POINTER(IMMDeviceCollection)), 'ppDevices')),
		COMMETHOD([], HRESULT, 'GetDefaultAudioEndpoint',
			(['in'], DWORD, 'dataFlow'),
			(['in'], DWORD, 'role'),
			(['out'], POINTER(POINTER(IMMDevice)), 'ppDevices')),)
