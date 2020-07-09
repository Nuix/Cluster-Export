# Menu Title: Cluster Export
# Needs Case: true
# @version 2.0.0

VERBOSE = false

begin
  # Nx Bootstrap
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
require 'rexml/document'
require 'csv'

ITEM_UTILITY = $utilities.get_item_utility

# Returns cluster name.
# Handles pseudoclusters with negative IDs.
#
# @param id [Integer] a cluster's ID
# @return [String] ID, or name for negative ID pseudoclusters
def id_to_name(id)
  return 'unclusterable' if id == -1
  return 'ignorable' if id == -2

  id.to_s
end

# Class for exporting items by cluster.
# * +@settings+ is input from the dialog
# * +@exporter+ is the BinaryExporter
# * +@target_directory+ is the export directory
# * +@failures+ is {GUID => error message} for items that failed to export
# * +@dialog+ is an Nx ProcessDialog
# * +@csv+ is the report CSV
class ClusterExport
  # Export items by cluster.
  #
  # @param settings [Hash] input from CustomExportSettings
  def initialize(settings)
    @settings = settings
    @exporter = $utilities.get_binary_exporter
    # strip the cluster run name in case it ends with a space
    @target_directory = File.join(@settings['dir'], @settings['cluster_run'].strip)
    # make sure output directory exists
    java.io.File.new(@target_directory).mkdirs
    @failures = {}
    ProgressDialog.forBlock do |progress_dialog|
      @dialog = initalize_dialog(progress_dialog, 'Cluster Export')
      @dialog.logMessage("Exporting to #{@target_directory}")
      @csv = report_csv_initialize
      run
      close
    end
  end

  # Closes the CSV, reports errors, and completes the dialog, or logs the abortion.
  def close
    @csv.close
    return @dialog.setMainStatusAndLogIt('Aborted') if @dialog.abortWasRequested

    report_errors unless @failures.empty?
    @dialog.setCompleted
  end

  # Exports the item and writes record to report.
  #
  # @param item [Item]
  # @param path [String]
  # @param cluster_name [String]
  def export_item(item, path, cluster_name)
    @dialog.logMessage("Exporting #{path}") if VERBOSE
    begin
      @exporter.exportItem(item, path)
    rescue StandardError => e
      message = e.message.to_s
      @dialog.logMessage(message)
      @failures[item.get_guid] = message
      path = 'ERROR'
    end
    # Write record to CSV
    @csv << report_csv_record(item, cluster_name, path)
  end

  # Gets deduped items from the cluster.
  #
  # @param cluster [Cluster]
  # @return [Set<Item>] deduped items
  def get_deduped_items(cluster)
    cluster_items = cluster.get_items.map(&:get_item)
    @dialog.logMessage("Cluster has #{cluster_items.size} items") if VERBOSE
    deduped_items = ITEM_UTILITY.deduplicate(cluster_items)
    if VERBOSE
      @dialog.logMessage("#{deduped_items.size} items after deduplicating")
    end
    @dialog.setSubProgress(0, deduped_items.size)
    deduped_items
  end

  # Calculates the export filenames up front.
  # Colliding filenames are be renamed to include MD5.
  #
  # @param items [Set<Item>] a cluster's deuped items
  # @return [{Item => String}] hash of items to their filename for export
  def get_export_details(items)
    export_details = {}
    # Track if any filenames occur more than once
    name_counts = Hash.new { |h, k| h[k] = 0 }
    items.each do |item|
      name = new_file_name(item)
      name_counts[name] += 1
      export_details[item] = name
    end
    # Fix filename collisions
    export_details.each do |i, n|
      if name_counts[n] > 1
        # insert MD5 before extension
        export_details[i] = String.new(n).insert((-1 - File.extname(n).size), " (#{i.getDigests.getMd5})")
      end
    end
    export_details
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

  # Returns true/false if the item is nil or empty.
  #
  # @param obj [Object]
  # @return [true, false] if the item is nil or empty (after strip)
  def nil_empty(obj)
    obj.nil? || obj.strip.empty?
  end

  # Initalizes CSV report.
  #
  # @return [CSV] with headers written
  def report_csv_initialize
    csv = CSV.open(File.join(@target_directory, 'Report.csv'), 'w:utf-8')
    # Add CSV headers
    csv << [
      'Item GUID',
      'Item Name',
      'Cluster ID',
      'Cluster Thread',
      'Cluster Endpoint Status',
      'MD5 Digest',
      'Original Path',
      'Tags',
      'Export Path'
    ]
    csv
  end

  # Generates the row for CSV report.
  #
  # @param item [Item]
  # @param cluster_name [String]
  # @param path [String]
  # @return [Array<String>] record for CSV
  def report_csv_record(item, cluster_name, path)
    id = "#{@settings['cluster_run']}-#{cluster_name}"
    [
      item.get_guid,
      item.get_localised_name,
      id,
      item.get_cluster_thread_indexes[id],
      item.get_cluster_endpoint_status[id],
      item.get_digests.get_md5,
      File.join(item.get_localised_path_names.to_a),
      item.getTags.join('; '),
      path
    ]
  end

  # Generates Errors.csv from items that failed to export (i.e. @failures).
  def report_errors
    @dialog.logMessage("#{@failures.size} items failed to export")
    errors_path = File.join(@target_directory, 'Errors.csv')
    @dialog.setSubStatusAndLogIt("Generating #{errors_path}")
    error_file = CSV.open(errors_path, 'w:utf-8')
    error_file << ['Item GUID', 'Error Message']
    @failures.each { |guid, msg| error_file << [guid, msg] }
    error_file.close
  end

  # Exports each cluster.
  def run
    clusters = @settings['clusters']
    @dialog.setMainStatusAndLogIt("Exporting #{clusters.size} clusters from #{@settings['cluster_run']}")
    @dialog.setMainProgress(0, clusters.size)
    clusters.each_with_index do |c, c_index|
      @dialog.setMainProgress(c_index)
      run_cluster(c)
      return nil if @dialog.abortWasRequested
    end
  end

  # Exports items from cluster.
  #
  # @param cluster [Cluster]
  def run_cluster(cluster)
    cluster_name = id_to_name(cluster.getId)
    @dialog.setSubStatusAndLogIt("Exporting Cluster #{cluster_name}")
    # Make sure output directory for cluster exists
    target_path = File.join(@target_directory, cluster_name)
    java.io.File.new(target_path).mkdirs
    # Calculate export details and check for filename collisions
    export_details = get_export_details(get_deduped_items(cluster))
    export_details.each_with_index do |(item, name), i_index|
      @dialog.setSubProgress(i_index)
      return nil if @dialog.abortWasRequested

      export_item(item, File.join(target_path, name), cluster_name)
    end
  end

  # Returns string santizied of illegal file system characters.
  #
  # @param value [String]
  # @return [String]
  def sanitize(value)
    value.gsub(%r{[\p{cc}<>:\"|?*\[\]\(\)\\/\t]}, '_')
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
    dedupe_count_cache = {}
    @main_tab.appendDynamicTable(identifier, 'Clusters', header, cluster_records(@cluster_runs[0].get_name)) do |record, column_index, _setting_value, _value|
      items = record.get_items
      case column_index
      when 0
        run_control.getSelectedItem
      when 1
        id_to_name(record.getId)
      when 2
        items.size
      when 3
        dedupe_count = dedupe_count_cache[record.getId]
        if dedupe_count.nil?
          dedupe_count = dedupe_count_cache[record.getId] = ITEM_UTILITY.deduplicate(items.map(&:get_item)).size
        end
        dedupe_count
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
    if values['dir'].empty?
      return CommonDialogs.showWarning('Please select export directory')
    end
    if values['clusters'].empty?
      return CommonDialogs.showWarning('Please select clusters')
    end

    true
  end
end

begin
  settings = ClusterExportSettings.new.input
  ClusterExport.new(settings) unless settings.nil?
end
