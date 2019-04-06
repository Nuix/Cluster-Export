# Menu Title: Cluster Export
# Needs Case: true
# @version 1.0.0

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

# Handles pseudoclusters with negative IDs.
#
# @param cluster [Cluster] the cluster
# @return [Integer, String] ID number, or name for negative ID pseudoclusters
def cluster_id(cluster)
  id = cluster.getId
  return 'unclusterable' if id == -1
  return 'ignorable' if id == -2

  id
end

ITEM_UTILITY = $utilities.get_item_utility

# Class for exporting items by cluster.
# * +@settings+ is input from the dialog
# * +@exporter+ is the BinaryExporter
# * +@target_directory+ is the export directory
# * +@exported+ is a hash of the exported files { path => [MD5]}
# * +@@dialog+ is an Nx ProcessDialog
class ClusterExport
  # Export items by cluster.
  #
  # @param settings [Hash] input from CustomExportSettings
  def initialize(settings)
    @settings = settings
    @exporter = $utilities.get_binary_exporter
    @target_directory = File.join(@settings['dir'], @settings['cluster_run'])
    @exported = Hash.new { |h, k| h[k] = [] }
    ProgressDialog.forBlock do |progress_dialog|
      @dialog = initalize_dialog(progress_dialog, 'Cluster Export')
      run
      close_nx
    end
  end

  # Completes the dialog, or logs the abortion.
  def close_nx
    return @dialog.setCompleted unless @dialog.abortWasRequested

    @dialog.setMainStatusAndLogIt('Aborted')
  end

  # Exports item.
  #
  # @param item [item]
  # @param target_directory_cluster [String] target directory for item
  def export_item(item, target_directory_cluster)
    target_path = File.join(target_directory_cluster, new_file_name(item))
    md5 = item.get_digests.get_md5
    @exported[target_path] << md5 # add before target_path changes
    target_path = new_file_path(target_path, md5) if File.exist?(target_path)
    @dialog.logMessage("Exporting #{target_path}")
    @exporter.exportItem(item, target_path)
  end

  # Initializes ProgressDialog
  #
  # @param progress_dialog [ProgressDialog]
  # @param title [String]
  # @return [ProgressDialog]
  def initalize_dialog(progress_dialog, title)
    progress_dialog.setTitle(title)
    progress_dialog.setLogVisible(true)
    progress_dialog.setTimestampLoggedMessages(true)
    progress_dialog
  end

  # Returns approriate file extension for item.
  #
  # @param item [item]
  # @return [String] file extension
  def new_file_ext(item)
    ext = item.getOriginalExtension
    return ext unless nil_empty(ext)

    ext = item.getCorrectedExtension
    return ext unless nil_empty(ext)

    ext = item.getType.getPreferredExtension
    return ext unless nil_empty(ext)

    'BIN'
  end

  # Returns sanitized file name for export.
  #
  # @param item [item]
  # @return [String] file name
  def new_file_name(item)
    sanitized_item_name = sanitize(item.getLocalisedName)
    # File data should have extension in its item name
    return sanitized_item_name if item.isFileData

    "#{sanitized_item_name}\.#{new_file_ext(item)}"
  end

  # Returns file path that includes digest.
  #
  # @param path [String]
  # @param digest [String]
  # @return [String] file name
  def new_file_path(path, digest)
    String.new(path).insert((-1 - File.extname(path).size), " (#{digest})")
  end

  # Returns true/false if the item is nil or empty.
  #
  # @param obj [Object]
  # @return [true, false] if the item is nil or empty (after strip)
  def nil_empty(obj)
    obj.nil? || obj.strip.empty?
  end

  # Renames iniitial files with name colisions to include digest.
  def rename_exported
    @dialog.setSubStatusAndLogIt('Including MD5 for filename colisions')
    paths = @exported.reject { |_k, v| v.size == 1 }
    @dialog.logMessage("Adding hash to #{paths.size} items")
    @dialog.setSubProgress(0, paths.size)
    paths.each_with_index do |(path, hashes), index|
      @dialog.setSubProgress(index)
      new_path = new_file_path(path, hashes.first)
      @dialog.logMessage("Renaming to #{new_path}")
      File.rename(path, new_path)
    end
  end

  # Exports each cluster.
  def run
    clusters = @settings['clusters']
    @dialog.setMainStatusAndLogIt("Exporting #{clusters.size} clusters from #{@settings['cluster_run']}")
    @dialog.logMessage("Target directory is #{@target_directory}")
    @dialog.setMainProgress(0, clusters.size)
    # Iterate each cluster
    clusters.each_with_index do |c, c_index|
      @dialog.setMainProgress(c_index)
      run_cluster(c)
      return nil if @dialog.abortWasRequested
    end
    # Rename files that had name colisions
    rename_exported
  end

  # Exports items from cluster.
  #
  # @param cluster [Cluster]
  def run_cluster(cluster)
    name = cluster_id(cluster)
    @dialog.setSubStatusAndLogIt("Exporting Cluster #{name}")
    # Iterate each item in cluster after deduplication
    items = cluster.get_items.map(&:get_item)
    @dialog.logMessage("Cluster has #{items.size} items")
    run_items(items, File.join(@target_directory, name.to_s))
  end

  # Exports items after deduplicating.
  #
  # @param items [Collection<Items>]
  # @param target_directory_cluster [String] export path
  def run_items(items, target_directory_cluster)
    # Make sure output directory exists
    FileUtils.mkdir_p(target_directory_cluster)
    deduped_items = ITEM_UTILITY.deduplicate(items)
    @dialog.logMessage("#{deduped_items.size} items after deduplicating")
    @dialog.setSubProgress(0, deduped_items.size)
    deduped_items.each_with_index do |i, i_index|
      @dialog.setSubProgress(i_index)
      export_item(i, target_directory_cluster)
      return nil if @dialog.abortWasRequested
    end
  end

  # Returns string santizied of illegal file system characters.
  #
  # @param value [String]
  # @return [String]
  def sanitize(value)
    value.gsub(%r{[<>:\"|?*\[\]\(\)\\/\t]}, '_')
  end
end

# Class for settings dialog.
# * +@cluster_runs+ are the cluster runs from the case
# * +@dialog+ is the TabbedCustomDialog
# * +@main_tab+ is the main tab
class ClusterExportSettings
  # Initialize dialog.
  def initialize
    @cluster_runs = $current_case.get_cluster_runs
    @dialog = TabbedCustomDialog.new('Cluster Export')
    @main_tab = @dialog.addTab('main_tab', 'Clusters')
    initialize_controls('dir', 'cluster_run', 'clusters')
    @dialog.validateBeforeClosing { |v| validate_input(v) }
  end

  # Display dialog and get input.
  #
  # @return [Hash] of input
  def input
    @dialog.display
    return nil if @dialog.getDialogResult == false

    @dialog.toMap
  end

  protected

  # Appends dynamic table to main tab.
  #
  # @param identifier [String] identifier for table
  # @param control [String] identifier for control that sets table records
  def append_dynamic_table(identifier, control)
    header = ['Cluster Run', 'ID', 'Items', 'Deduplicated Items']
    run_control = @main_tab.getControl(control)
    @main_tab.appendDynamicTable(identifier, 'Clusters', header, cluster_records(@cluster_runs[0].get_name)) do |record, column_index, _setting_value, _value|
      items = record.get_items
      case column_index
      when 0
        run_control.getSelectedItem
      when 1
        cluster_id(record)
      when 2
        items.size
      when 3
        ITEM_UTILITY.deduplicate(items.map(&:get_item)).size
      end
    end
    initialize_dynamic_table(identifier, run_control)
  end

  # Returns array of clusters sorted by ID.
  #
  # @param name [String] cluster run name
  # @return [Array] of clusters
  def cluster_records(name)
    run = @cluster_runs.find { |r| r.get_name == name }
    return nil if run.nil?

    run.get_clusters.to_a.sort_by!(&:get_id)
  end

  # Appends controls to main tab.
  #
  # @param path [String] identifier for export directory
  # @param combo [String] identifier for combo box
  # @param table [String] identifier for table
  def initialize_controls(path, combo, table)
    @main_tab.appendDirectoryChooser(path, 'Export Directory')
    choices = @cluster_runs.map(&:get_name)
    @main_tab.appendComboBox(combo, 'Cluster Run', choices)
    append_dynamic_table(table, combo)
  end

  # Initializes listener to set records and check status in table.
  #
  # @param table [String] identifier for the table
  # @param run_control [java.awt.Component] Java Swing control the sets table
  def initialize_dynamic_table(table, run_control)
    table_model = @main_tab.getControl(table).getModel
    initialize_table_checks(table_model)
    run_control.addActionListener do |_e|
      table_model.uncheckDisplayedRecords
      table_model.setRecords(cluster_records(run_control.getSelectedItem))
      initialize_table_checks(table_model)
    end
  end

  # Checks all records in table, then uncheck pseudoclusters.
  #
  # @param table_model [DynamicTableModel]
  def initialize_table_checks(table_model)
    table_model.checkDisplayedRecords
    (0..1).each do |row|
      break unless table_model.getValueAt(row, 2).is_a?(String)

      table_model.setCheckedAtIndex(row, false)
    end
  end

  # Validation function for input.
  #
  # @param values [Hash] input values
  # @return [true, false] if in valid run state
  def validate_input(values)
    return CommonDialogs.showWarning('Please select export directory') if values['dir'].empty?
    return CommonDialogs.showWarning('Please select clusters') if values['clusters'].empty?

    true
  end
end

begin
  settings = ClusterExportSettings.new.input
  ClusterExport.new(settings) unless settings.nil?
end
