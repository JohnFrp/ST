import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import '../models/stock.dart';
import '../database/database_helper.dart';

class ImportScreen extends StatefulWidget {
  const ImportScreen({super.key});

  @override
  State<ImportScreen> createState() => _ImportScreenState();
}

class _ImportScreenState extends State<ImportScreen> {
  final DatabaseHelper _databaseHelper = DatabaseHelper();
  bool _isImporting = false;
  String _importResult = '';
  String? _selectedFileName;
  int _totalRows = 0;
  int _processedRows = 0;
  List<String> _debugInfo = [];

  Future<void> _pickAndImportExcel() async {
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
        allowMultiple: false,
      );

      if (result == null || result.files.isEmpty) {
        _showError('No file selected');
        return;
      }

      final file = result.files.first;

      // Debug file information
      _addDebugInfo('File selected: ${file.name}');
      _addDebugInfo('File size: ${file.size} bytes');
      _addDebugInfo('File extension: ${file.extension}');
      _addDebugInfo('File path: ${file.path}');
      _addDebugInfo('File bytes: ${file.bytes?.length} bytes');

      // Validate file
      if (!file.name.toLowerCase().endsWith('.xlsx') &&
          !file.name.toLowerCase().endsWith('.xls')) {
        _showError('Please select a valid Excel file (.xlsx or .xls)');
        return;
      }

      setState(() {
        _isImporting = true;
        _importResult = '';
        _selectedFileName = file.name;
        _totalRows = 0;
        _processedRows = 0;
        _debugInfo.clear();
      });

      // Get file bytes with fallback methods
      Uint8List? fileBytes = await _getFileBytes(file);

      if (fileBytes == null || fileBytes.isEmpty) {
        throw Exception(
            'Could not read file content. File may be corrupted or inaccessible.');
      }

      _addDebugInfo('Successfully read file bytes: ${fileBytes.length} bytes');

      // Parse and import Excel data
      await _importExcelData(fileBytes);
    } catch (e) {
      _showError('Error importing file: $e');
      _addDebugInfo('Error: $e');
    } finally {
      setState(() {
        _isImporting = false;
      });
    }
  }

  Future<Uint8List?> _getFileBytes(PlatformFile file) async {
    // Method 1: Use bytes from file picker (primary method)
    if (file.bytes != null && file.bytes!.isNotEmpty) {
      _addDebugInfo('Using bytes from file picker');
      return file.bytes;
    }

    // Method 2: Read from file path (fallback method)
    if (file.path != null) {
      try {
        _addDebugInfo('Reading from file path: ${file.path}');
        final fileObject = File(file.path!);
        if (await fileObject.exists()) {
          final bytes = await fileObject.readAsBytes();
          _addDebugInfo(
              'Successfully read ${bytes.length} bytes from file path');
          return bytes;
        } else {
          _addDebugInfo('File does not exist at path: ${file.path}');
        }
      } catch (e) {
        _addDebugInfo('Error reading from file path: $e');
      }
    }

    // Method 3: Try alternative file picker with different settings
    _addDebugInfo('Trying alternative file picker method...');
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
        allowMultiple: false,
        withData: true, // Force reading file data
        withReadStream: false,
      );

      if (result != null && result.files.isNotEmpty) {
        final alternativeFile = result.files.first;
        if (alternativeFile.bytes != null &&
            alternativeFile.bytes!.isNotEmpty) {
          _addDebugInfo(
              'Alternative method successful: ${alternativeFile.bytes!.length} bytes');
          return alternativeFile.bytes;
        }
      }
    } catch (e) {
      _addDebugInfo('Alternative file picker failed: $e');
    }

    _addDebugInfo('All methods failed to read file bytes');
    return null;
  }

  Future<void> _importExcelData(Uint8List bytes) async {
    try {
      _addDebugInfo('Starting Excel parsing...');

      // Decode Excel file
      final excel = Excel.decodeBytes(bytes);

      _addDebugInfo('Excel tables found: ${excel.tables.keys.length}');
      _addDebugInfo('Table names: ${excel.tables.keys.join(', ')}');

      if (excel.tables.isEmpty) {
        throw Exception('No sheets found in Excel file');
      }

      // Get the first sheet
      final sheetName = excel.tables.keys.first;
      final table = excel.tables[sheetName]!;

      _addDebugInfo('Using sheet: $sheetName');
      _addDebugInfo('Total rows in sheet: ${table.rows.length}');

      if (table.rows.isEmpty) {
        throw Exception('Excel sheet is empty');
      }

      // Get header row and find column indices
      final headers = _getHeaders(table.rows[0]);
      _addDebugInfo(
          'Headers found: ${headers.where((h) => h.isNotEmpty).join(', ')}');

      // Validate required columns
      final validationResult = _validateHeaders(headers);
      if (!validationResult.isValid) {
        throw Exception(validationResult.errorMessage);
      }

      setState(() {
        _totalRows = table.rows.length - 1; // Exclude header row
      });

      _addDebugInfo('Data rows to process: $_totalRows');

      // Process data rows
      int stocksAdded = 0;
      int stocksUpdated = 0;
      int errors = 0;
      List<String> processedItems = [];

      for (int i = 1; i < table.rows.length; i++) {
        try {
          final row = table.rows[i];

          if (_isRowEmpty(row)) {
            _addDebugInfo('Row ${i + 1} is empty, skipping');
            continue;
          }

          _addDebugInfo('Processing row ${i + 1}: ${_rowToString(row)}');

          final stock = _parseStockFromRow(row, headers);

          if (stock != null) {
            _addDebugInfo(
                'Parsed stock: ${stock.name} - Buy: ${stock.buyPrice}, Sell: ${stock.sellPrice}, Qty: ${stock.stockQuantity}');

            // Check if stock already exists
            final existingStocks =
                await _databaseHelper.searchStocks(stock.name);

            if (existingStocks.isNotEmpty) {
              // Update existing stock
              final existing = existingStocks.first;
              _addDebugInfo('Updating existing stock: ${existing.name}');

              existing
                ..genericName = stock.genericName
                ..manufacturer = stock.manufacturer
                ..category = stock.category
                ..buyPrice = stock.buyPrice
                ..sellPrice = stock.sellPrice
                ..stockQuantity = stock.stockQuantity
                ..expiryDate = stock.expiryDate;

              await _databaseHelper.updateStock(existing);
              stocksUpdated++;
              processedItems.add('🔄 UPDATED: ${stock.name}');
            } else {
              // Add new stock
              _addDebugInfo('Adding new stock: ${stock.name}');
              await _databaseHelper.insertStock(stock);
              stocksAdded++;
              processedItems.add('✅ ADDED: ${stock.name}');
            }
          } else {
            _addDebugInfo('Row ${i + 1} could not be parsed as stock');
            errors++;
          }

          setState(() {
            _processedRows = i;
          });
        } catch (e) {
          errors++;
          _addDebugInfo('❌ Error processing row ${i + 1}: $e');
          processedItems.add('❌ ERROR Row ${i + 1}: $e');
        }
      }

      // Build final result
      final resultBuffer = StringBuffer();
      resultBuffer.writeln('Import completed!\n');
      resultBuffer.writeln('✅ $stocksAdded new items added');
      resultBuffer.writeln('🔄 $stocksUpdated existing items updated');
      resultBuffer.writeln('❌ $errors errors encountered');
      resultBuffer.writeln('📊 Total rows processed: $_totalRows\n');

      // Add first few processed items for preview
      if (processedItems.isNotEmpty) {
        resultBuffer.writeln('Processed items:');
        final previewItems = processedItems.take(10).toList();
        for (final item in previewItems) {
          resultBuffer.writeln('• $item');
        }
        if (processedItems.length > 10) {
          resultBuffer.writeln('• ... and ${processedItems.length - 10} more');
        }
      }

      setState(() {
        _importResult = resultBuffer.toString();
      });

      _showSuccess('Import completed successfully!');
    } catch (e) {
      _addDebugInfo('Fatal error: $e');
      rethrow;
    }
  }

  List<String> _getHeaders(List<Data?> headerRow) {
    final headers = <String>[];
    for (int i = 0; i < headerRow.length; i++) {
      final cell = headerRow[i];
      final value = cell?.value?.toString().trim() ?? '';
      headers.add(value.toLowerCase());
      _addDebugInfo('Header column $i: "$value"');
    }
    return headers;
  }

  HeaderValidationResult _validateHeaders(List<String> headers) {
    final requiredColumns = [
      'name',
      'buy_price',
      'sell_price',
      'stock_quantity'
    ];
    final missingColumns =
        requiredColumns.where((col) => !headers.contains(col)).toList();

    _addDebugInfo(
        'Looking for required columns: ${requiredColumns.join(', ')}');
    _addDebugInfo('Missing columns: ${missingColumns.join(', ')}');

    if (missingColumns.isNotEmpty) {
      return HeaderValidationResult(
        isValid: false,
        errorMessage:
            'Missing required columns: ${missingColumns.join(', ')}\n\n'
            'Required columns: name, buy_price, sell_price, stock_quantity\n'
            'Found columns: ${headers.where((h) => h.isNotEmpty).join(', ')}',
      );
    }

    return HeaderValidationResult(isValid: true);
  }

  bool _isRowEmpty(List<Data?> row) {
    if (row.isEmpty) return true;

    // Check if first cell is empty
    final firstCell = row[0]?.value;
    if (firstCell == null || firstCell.toString().trim().isEmpty) {
      return true;
    }

    return false;
  }

  String _rowToString(List<Data?> row) {
    return row.map((cell) => cell?.value?.toString() ?? 'null').join(' | ');
  }

  Stock? _parseStockFromRow(List<Data?> row, List<String> headers) {
    try {
      final name = _getCellValue(row, headers, 'name');
      _addDebugInfo('Name value: "$name"');

      if (name == null || name.isEmpty) {
        _addDebugInfo('Skipping row - no name found');
        return null; // Skip rows without name
      }

      // Safe numeric parsing with fallbacks
      double safeParseDouble(String? value) {
        if (value == null || value.isEmpty) return 0.0;
        final cleanValue = value.replaceAll(',', '');
        return double.tryParse(cleanValue) ?? 0.0;
      }

      int safeParseInt(String? value) {
        if (value == null || value.isEmpty) return 0;
        final cleanValue = value.replaceAll(',', '');
        return int.tryParse(cleanValue) ?? 0;
      }

      final buyPrice =
          safeParseDouble(_getCellValue(row, headers, 'buy_price'));
      final sellPrice =
          safeParseDouble(_getCellValue(row, headers, 'sell_price'));
      final stockQuantity =
          safeParseInt(_getCellValue(row, headers, 'stock_quantity'));

      _addDebugInfo(
          'Prices - Buy: $buyPrice, Sell: $sellPrice, Qty: $stockQuantity');

      // Validate that we have at least some data
      if (buyPrice == 0.0 && sellPrice == 0.0 && stockQuantity == 0) {
        _addDebugInfo('All numeric fields are zero, skipping row');
        return null;
      }

      return Stock(
        name: name,
        genericName: _getCellValue(row, headers, 'generic_name') ?? '',
        manufacturer: _getCellValue(row, headers, 'manufacturer') ?? '',
        category: _getCellValue(row, headers, 'category') ?? '',
        buyPrice: buyPrice,
        sellPrice: sellPrice,
        stockQuantity: stockQuantity,
        expiryDate: _parseDate(_getCellValue(row, headers, 'expiry_date')),
      );
    } catch (e) {
      _addDebugInfo('Error parsing stock: $e');
      return null;
    }
  }

  String? _getCellValue(
      List<Data?> row, List<String> headers, String columnName) {
    final index = headers.indexOf(columnName);
    if (index == -1) {
      _addDebugInfo('Column "$columnName" not found in headers');
      return null;
    }
    if (index >= row.length) {
      _addDebugInfo(
          'Column index $index out of bounds (row length: ${row.length})');
      return null;
    }

    final cell = row[index];
    final value = cell?.value;
    final result = value?.toString().trim();

    _addDebugInfo(
        'Cell [$columnName] at index $index: "$result" (raw: "$value")');

    return result;
  }

  double? _getNumericValue(
      List<Data?> row, List<String> headers, String columnName) {
    final value = _getCellValue(row, headers, columnName);
    if (value == null || value.isEmpty) {
      _addDebugInfo('Numeric value for "$columnName" is null or empty');
      return null;
    }

    // Handle both integer and decimal numbers
    final cleanValue = value.replaceAll(',', '');

    // Try parsing as double first, then as int and convert to double
    double? numericValue = double.tryParse(cleanValue);
    if (numericValue == null) {
      final intValue = int.tryParse(cleanValue);
      if (intValue != null) {
        numericValue = intValue.toDouble();
      }
    }

    _addDebugInfo(
        'Numeric conversion: "$value" -> "$cleanValue" -> $numericValue');

    return numericValue;
  }

  DateTime? _parseDate(String? dateString) {
    if (dateString == null || dateString.isEmpty) {
      _addDebugInfo('Date string is null or empty');
      return null;
    }

    _addDebugInfo('Parsing date: "$dateString"');

    try {
      // Try parsing as DateTime first
      final dateTime = DateTime.tryParse(dateString);
      if (dateTime != null) {
        _addDebugInfo('Successfully parsed as DateTime: $dateTime');
        return dateTime;
      }

      // Try common date formats
      final formats = [
        'yyyy-MM-dd',
        'dd/MM/yyyy',
        'MM/dd/yyyy',
        'yyyy/MM/dd',
      ];

      for (final format in formats) {
        final parts = dateString.split(RegExp(r'[/-]'));
        if (parts.length == 3) {
          int? year, month, day;

          if (format == 'yyyy-MM-dd' || format == 'yyyy/MM/dd') {
            year = int.tryParse(parts[0]);
            month = int.tryParse(parts[1]);
            day = int.tryParse(parts[2]);
          } else if (format == 'dd/MM/yyyy') {
            day = int.tryParse(parts[0]);
            month = int.tryParse(parts[1]);
            year = int.tryParse(parts[2]);
          } else if (format == 'MM/dd/yyyy') {
            month = int.tryParse(parts[0]);
            day = int.tryParse(parts[1]);
            year = int.tryParse(parts[2]);
          }

          if (year != null && month != null && day != null) {
            // Handle 2-digit years
            if (year < 100) {
              year += 2000;
            }

            final result = DateTime(year, month, day);
            _addDebugInfo('Successfully parsed with format $format: $result');
            return result;
          }
        }
      }

      _addDebugInfo('Could not parse date: $dateString');
      return null;
    } catch (e) {
      _addDebugInfo('Error parsing date: $dateString - $e');
      return null;
    }
  }

  void _addDebugInfo(String message) {
    print('IMPORT DEBUG: $message');
    if (_debugInfo.length < 100) {
      // Limit to prevent memory issues
      _debugInfo.add('${DateTime.now().toString().split('.').last}: $message');
    }
  }

  void _showError(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(message),
        backgroundColor: Colors.red,
        duration: const Duration(seconds: 5),
      ),
    );
  }

  void _showSuccess(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(message),
        backgroundColor: Colors.green,
        duration: const Duration(seconds: 3),
      ),
    );
  }

  void _showDebugInfo() {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('Debug Information'),
        content: SizedBox(
          width: double.maxFinite,
          child: ListView.builder(
            shrinkWrap: true,
            itemCount: _debugInfo.length,
            itemBuilder: (context, index) {
              return Text(_debugInfo[index]);
            },
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('Close'),
          ),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Import Stock Data'),
        actions: [
          if (_debugInfo.isNotEmpty)
            IconButton(
              icon: const Icon(Icons.bug_report),
              onPressed: _showDebugInfo,
              tooltip: 'Show Debug Info',
            ),
        ],
      ),
      body: SingleChildScrollView( // FIX: Wrap in SingleChildScrollView
        padding: const EdgeInsets.all(16),
        child: Column(
          children: [
            // Import Card
            Card(
              elevation: 4,
              child: Padding(
                padding: const EdgeInsets.all(20),
                child: Column(
                  children: [
                    const Icon(Icons.file_upload, size: 64, color: Colors.blue),
                    const SizedBox(height: 16),
                    const Text(
                      'Import Excel File',
                      style:
                          TextStyle(fontSize: 22, fontWeight: FontWeight.bold),
                    ),
                    const SizedBox(height: 8),
                    Text(
                      _selectedFileName ?? 'No file selected',
                      style: TextStyle(
                        fontSize: 16,
                        color: _selectedFileName != null
                            ? Colors.green
                            : Colors.grey,
                        fontWeight: FontWeight.w500,
                      ),
                      textAlign: TextAlign.center,
                    ),
                    const SizedBox(height: 20),
                    if (_isImporting) ...[
                      _buildProgressIndicator(),
                      const SizedBox(height: 16),
                    ],
                    FilledButton.icon(
                      onPressed: _isImporting ? null : _pickAndImportExcel,
                      icon: const Icon(Icons.upload_file),
                      label: const Text('Choose Excel File'),
                      style: FilledButton.styleFrom(
                        padding: const EdgeInsets.symmetric(
                            horizontal: 24, vertical: 12),
                      ),
                    ),
                  ],
                ),
              ),
            ),

            const SizedBox(height: 20),

            // Import Result
            if (_importResult.isNotEmpty)
              Card(
                color: Colors.green[50],
                child: Padding(
                  padding: const EdgeInsets.all(16),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Row(
                        children: [
                          Icon(Icons.check_circle, color: Colors.green),
                          SizedBox(width: 8),
                          Text(
                            'Import Result',
                            style: TextStyle(
                                fontSize: 18,
                                fontWeight: FontWeight.bold,
                                color: Colors.green),
                          ),
                        ],
                      ),
                      const SizedBox(height: 8),
                      Text(
                        _importResult,
                        style: const TextStyle(fontSize: 14),
                      ),
                    ],
                  ),
                ),
              ),

            const SizedBox(height: 20), // FIX: Replaced Spacer with SizedBox

            // Danger Zone
            Card(
              color: Colors.red[50],
              elevation: 2,
              child: Padding(
                padding: const EdgeInsets.all(16),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Row(
                      children: [
                        Icon(Icons.warning, color: Colors.red),
                        SizedBox(width: 8),
                        Text(
                          'Danger Zone',
                          style: TextStyle(
                              fontSize: 18,
                              fontWeight: FontWeight.bold,
                              color: Colors.red),
                        ),
                      ],
                    ),
                    const SizedBox(height: 8),
                    const Text('Clear all stock data from the database'),
                    const SizedBox(height: 16),
                    FilledButton.icon(
                      style:
                          FilledButton.styleFrom(backgroundColor: Colors.red),
                      onPressed: _isImporting ? null : _clearAllStocks,
                      icon: const Icon(Icons.delete_forever),
                      label: const Text('Clear All Stock Data'),
                    ),
                  ],
                ),
              ),
            ),

            const SizedBox(height: 20), // Add some bottom padding
          ],
        ),
      ),
    );
  }

  Widget _buildProgressIndicator() {
    return Column(
      children: [
        LinearProgressIndicator(
          value: _totalRows > 0 ? _processedRows / _totalRows : 0,
          backgroundColor: Colors.grey[300],
          valueColor: AlwaysStoppedAnimation<Color>(Colors.blue),
        ),
        const SizedBox(height: 8),
        Text(
          'Processing... $_processedRows/$_totalRows',
          style: const TextStyle(color: Colors.blue),
        ),
      ],
    );
  }

  Future<void> _clearAllStocks() async {
    final shouldDelete = await showDialog<bool>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('Clear All Stock'),
        content: const Text(
            'This will permanently delete ALL stock data. This action cannot be undone.'),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(false),
            child: const Text('Cancel'),
          ),
          TextButton(
            onPressed: () => Navigator.of(context).pop(true),
            child: const Text('Clear All', style: TextStyle(color: Colors.red)),
          ),
        ],
      ),
    );

    if (shouldDelete == true) {
      try {
        await _databaseHelper.deleteAllStocks();
        setState(() {
          _importResult = 'All stock data cleared successfully';
        });
        _showSuccess('All stock data cleared successfully');
      } catch (e) {
        _showError('Error clearing stock: $e');
      }
    }
  }

  // [Keep all your helper methods unchanged...]
  void _addDebugInfo(String message) {
    print('IMPORT DEBUG: $message');
    if (_debugInfo.length < 100) {
      _debugInfo.add('${DateTime.now().toString().split('.').last}: $message');
    }
  }

  void _showError(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(message),
        backgroundColor: Colors.red,
        duration: const Duration(seconds: 5),
      ),
    );
  }

  void _showSuccess(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(message),
        backgroundColor: Colors.green,
        duration: const Duration(seconds: 3),
      ),
    );
  }

  void _showDebugInfo() {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: const Text('Debug Information'),
        content: SizedBox(
          width: double.maxFinite,
          child: ListView.builder(
            shrinkWrap: true,
            itemCount: _debugInfo.length,
            itemBuilder: (context, index) {
              return Text(_debugInfo[index]);
            },
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('Close'),
          ),
        ],
      ),
    );
  }

  // [Keep all your Excel parsing methods unchanged...]
  List<String> _getHeaders(List<Data?> headerRow) {
    final headers = <String>[];
    for (int i = 0; i < headerRow.length; i++) {
      final cell = headerRow[i];
      final value = cell?.value?.toString().trim() ?? '';
      headers.add(value.toLowerCase());
      _addDebugInfo('Header column $i: "$value"');
    }
    return headers;
  }

  HeaderValidationResult _validateHeaders(List<String> headers) {
    final requiredColumns = [
      'name',
      'buy_price',
      'sell_price',
      'stock_quantity'
    ];
    final missingColumns =
        requiredColumns.where((col) => !headers.contains(col)).toList();

    _addDebugInfo(
        'Looking for required columns: ${requiredColumns.join(', ')}');
    _addDebugInfo('Missing columns: ${missingColumns.join(', ')}');

    if (missingColumns.isNotEmpty) {
      return HeaderValidationResult(
        isValid: false,
        errorMessage:
            'Missing required columns: ${missingColumns.join(', ')}\n\n'
            'Required columns: name, buy_price, sell_price, stock_quantity\n'
            'Found columns: ${headers.where((h) => h.isNotEmpty).join(', ')}',
      );
    }

    return HeaderValidationResult(isValid: true);
  }

  bool _isRowEmpty(List<Data?> row) {
    if (row.isEmpty) return true;
    final firstCell = row[0]?.value;
    if (firstCell == null || firstCell.toString().trim().isEmpty) {
      return true;
    }
    return false;
  }

  String _rowToString(List<Data?> row) {
    return row.map((cell) => cell?.value?.toString() ?? 'null').join(' | ');
  }

  Stock? _parseStockFromRow(List<Data?> row, List<String> headers) {
    try {
      final name = _getCellValue(row, headers, 'name');
      _addDebugInfo('Name value: "$name"');

      if (name == null || name.isEmpty) {
        _addDebugInfo('Skipping row - no name found');
        return null;
      }

      double safeParseDouble(String? value) {
        if (value == null || value.isEmpty) return 0.0;
        final cleanValue = value.replaceAll(',', '');
        return double.tryParse(cleanValue) ?? 0.0;
      }

      int safeParseInt(String? value) {
        if (value == null || value.isEmpty) return 0;
        final cleanValue = value.replaceAll(',', '');
        return int.tryParse(cleanValue) ?? 0;
      }

      final buyPrice =
          safeParseDouble(_getCellValue(row, headers, 'buy_price'));
      final sellPrice =
          safeParseDouble(_getCellValue(row, headers, 'sell_price'));
      final stockQuantity =
          safeParseInt(_getCellValue(row, headers, 'stock_quantity'));

      _addDebugInfo(
          'Prices - Buy: $buyPrice, Sell: $sellPrice, Qty: $stockQuantity');

      if (buyPrice == 0.0 && sellPrice == 0.0 && stockQuantity == 0) {
        _addDebugInfo('All numeric fields are zero, skipping row');
        return null;
      }

      return Stock(
        name: name,
        genericName: _getCellValue(row, headers, 'generic_name') ?? '',
        manufacturer: _getCellValue(row, headers, 'manufacturer') ?? '',
        category: _getCellValue(row, headers, 'category') ?? '',
        buyPrice: buyPrice,
        sellPrice: sellPrice,
        stockQuantity: stockQuantity,
        expiryDate: _parseDate(_getCellValue(row, headers, 'expiry_date')),
      );
    } catch (e) {
      _addDebugInfo('Error parsing stock: $e');
      return null;
    }
  }

  String? _getCellValue(
      List<Data?> row, List<String> headers, String columnName) {
    final index = headers.indexOf(columnName);
    if (index == -1) {
      _addDebugInfo('Column "$columnName" not found in headers');
      return null;
    }
    if (index >= row.length) {
      _addDebugInfo(
          'Column index $index out of bounds (row length: ${row.length})');
      return null;
    }

    final cell = row[index];
    final value = cell?.value;
    final result = value?.toString().trim();

    _addDebugInfo(
        'Cell [$columnName] at index $index: "$result" (raw: "$value")');

    return result;
  }

  DateTime? _parseDate(String? dateString) {
    if (dateString == null || dateString.isEmpty) {
      _addDebugInfo('Date string is null or empty');
      return null;
    }

    _addDebugInfo('Parsing date: "$dateString"');

    try {
      final dateTime = DateTime.tryParse(dateString);
      if (dateTime != null) {
        _addDebugInfo('Successfully parsed as DateTime: $dateTime');
        return dateTime;
      }

      final formats = [
        'yyyy-MM-dd',
        'dd/MM/yyyy',
        'MM/dd/yyyy',
        'yyyy/MM/dd',
      ];

      for (final format in formats) {
        final parts = dateString.split(RegExp(r'[/-]'));
        if (parts.length == 3) {
          int? year, month, day;

          if (format == 'yyyy-MM-dd' || format == 'yyyy/MM/dd') {
            year = int.tryParse(parts[0]);
            month = int.tryParse(parts[1]);
            day = int.tryParse(parts[2]);
          } else if (format == 'dd/MM/yyyy') {
            day = int.tryParse(parts[0]);
            month = int.tryParse(parts[1]);
            year = int.tryParse(parts[2]);
          } else if (format == 'MM/dd/yyyy') {
            month = int.tryParse(parts[0]);
            day = int.tryParse(parts[1]);
            year = int.tryParse(parts[2]);
          }

          if (year != null && month != null && day != null) {
            if (year < 100) {
              year += 2000;
            }
            final result = DateTime(year, month, day);
            _addDebugInfo('Successfully parsed with format $format: $result');
            return result;
          }
        }
      }

      _addDebugInfo('Could not parse date: $dateString');
      return null;
    } catch (e) {
      _addDebugInfo('Error parsing date: $dateString - $e');
      return null;
    }
  }
}

class HeaderValidationResult {
  final bool isValid;
  final String? errorMessage;

  HeaderValidationResult({
    required this.isValid,
    this.errorMessage,
  });
}