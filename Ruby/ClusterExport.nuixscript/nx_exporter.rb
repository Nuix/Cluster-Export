# Base class for exporting items with Nx.
# @version 2.1.0

begin # Nx Bootstrap
  require File.join(__dir__, 'Nx.jar')
  java_import 'com.nuix.nx.NuixConnection'
  java_import 'com.nuix.nx.LookAndFeelHelper'
  java_import 'com.nuix.nx.dialogs.ChoiceDialog'
  java_import 'com.nuix.nx.dialogs.CommonDialogs'
  java_import 'com.nuix.nx.dialogs.ProcessingStatusDialog'
  java_import 'com.nuix.nx.dialogs.ProgressDialog'
  java_import 'com.nuix.nx.dialogs.TabbedCustomDialog'
  java_import 'com.nuix.nx.digest.DigestHelper'
  java_import 'com.nuix.nx.controls.models.Choice'
  LookAndFeelHelper.setWindowsIfMetal
  NuixConnection.setUtilities($utilities)
  NuixConnection.setCurrentNuixVersion(NUIX_VERSION)
end
require 'fileutils'
require 'rexml/document'

# Class for Nx Dialog.
# * +@@dialog+ is an Nx ProcessDialog
# * +@@progress+ represents the main progress
class NxExporter
  # Initializes Progress Dialog.
  #
  # @param progress_dialog [ProgressDialog] the progress dialog
  # @param title [String] the title
  def initialize(progress_dialog, title)
    progress_dialog.setTitle(title)
    progress_dialog.setLogVisible(true)
    progress_dialog.setTimestampLoggedMessages(true)
    @@dialog = progress_dialog
    @@progress = 0
  end

  # Prompts user to choose a directory.
  #
  # @return [String, nil] absolute path or nil if canceled
  def self.choose_export_dir
    java_import javax.swing.JFileChooser
    c = JFileChooser.new
    c.setFileSelectionMode(JFileChooser::DIRECTORIES_ONLY)
    c.setDialogTitle('Select Export Directory')
    return nil unless c.showOpenDialog(nil) == JFileChooser::APPROVE_OPTION

    c.getSelectedFile.getAbsolutePath
  end

  # Increments and sets main progress.
  def advance_main
    @@progress += 1
    @@dialog.setMainProgress(@@progress)
  end

  # Completes the dialog, or logs the abortion.
  def close_nx
    return @@dialog.setCompleted unless @@dialog.abortWasRequested

    @@dialog.setMainStatusAndLogIt('Aborted')
  end
end
