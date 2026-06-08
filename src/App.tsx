import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { Search, CheckCircle, XCircle, AlertTriangle, FileSpreadsheet, Trash2, Settings, BarChart2, ScanLine, Ban, ClipboardList, X, Download, Save, UploadCloud, Loader2, ExternalLink } from 'lucide-react';

// ----------------------------------------------------------------------
// HELPER FUNCTIONS (Pure Functions)
// ----------------------------------------------------------------------
const getRemark = (row) => {
  if (!row) return "";
  let remarks = [];
  if (String(row["KMH"]) === "1") remarks.push("KMH");
  if (String(row["Demo"]) === "1") remarks.push("Demo");
  return remarks.join(", ");
};

const createRecord = (rawScan, status, rowData = null, extraData = {}) => {
  return {
    id: Date.now() + Math.random(),
    rawScan: String(rawScan),
    status,
    sp: rowData ? String(rowData["SP"] || "").trim() : '', 
    scCode: rowData ? String(rowData["SC Code"] || rowData["SC code"] || "").trim() : '', 
    ttbh: rowData ? String(rowData["Warehouse Name"]).trim() : '',
    soRO: rowData ? rowData["After-sales work order No."] : '',
    maLK: rowData ? rowData["Defective material code"] : '',
    tenLK: rowData ? rowData["Product Name"] : '',
    model: rowData ? (rowData["Product Model"] || rowData["Model"] || '') : '', 
    phanLoai: rowData ? rowData["Type"] : '',
    bhDv: rowData ? rowData["Repair Type"] : '',
    slg: rowData ? (rowData["Consumed quantity"] || rowData["Consumed"] || '') : '', 
    remark: getRemark(rowData),
    ...extraData
  };
};

const renderProgressBar = (scannedCount, totalCount, label, type) => {
  const percent = totalCount > 0 ? Math.round((scannedCount / totalCount) * 100) : 0;
  const isIW = type === 'IW';
  const colorBg = isIW ? 'bg-emerald-100' : 'bg-blue-100';
  const colorFg = isIW ? 'bg-emerald-500' : 'bg-blue-500';
  const colorText = isIW ? 'text-emerald-700' : 'text-blue-700';

  return (
    <div className="w-full mb-2.5 last:mb-0">
      <div className="flex justify-between items-end mb-1">
        <span className={`font-bold text-[11px] uppercase tracking-wider ${colorText}`}>{label}</span>
        <span className={`font-bold text-xs ${colorText}`}>
          {percent}% <span className="text-[10px] opacity-70 font-medium ml-0.5">({scannedCount}/{totalCount})</span>
        </span>
      </div>
      <div className={`w-full h-2 rounded-full ${colorBg} overflow-hidden shadow-inner`}>
        <div className={`h-full rounded-full ${colorFg} transition-all duration-700 ease-out`} style={{ width: `${percent}%` }}></div>
      </div>
    </div>
  );
};

// Hàm tạo cấu trúc file Excel thống nhất giữa các chức năng xuất
const generateWorkbookData = (selectedWarehouse, currentDisplayedRecords, detailedProgressList) => {
  const workbook = window.XLSX.utils.book_new();

  const statsData = [
    {
      "Nhóm Phân Loại": "Bảo Hành (IW)",
      "Tổng Cần Scan": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW').length,
      "Đã Scan Khớp": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW' && r.isScanned).length,
      "Còn Thiếu": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW' && !r.isScanned).length,
    },
    {
      "Nhóm Phân Loại": "Dịch Vụ (OOW)",
      "Tổng Cần Scan": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW').length,
      "Đã Scan Khớp": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && r.isScanned).length,
      "Còn Thiếu": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !r.isScanned).length,
    },
    {
      "Nhóm Phân Loại": "Khác",
      "Tổng Cần Scan": detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase())).length,
      "Đã Scan Khớp": detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase()) && r.isScanned).length,
      "Còn Thiếu": detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase()) && !r.isScanned).length,
    }
  ];
  const wsStats = window.XLSX.utils.json_to_sheet(statsData);
  wsStats['!cols'] = [{ wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];
  window.XLSX.utils.book_append_sheet(workbook, wsStats, "ThongKe");

  // FORM CỘT TRÁI SANG PHẢI ĐỒNG NHẤT
  const exportScannedData = currentDisplayedRecords.map(record => ({
    "Trạng thái": record.status, 
    "Cột SP": record.sp || record.rawScan, 
    "SC Code": record.scCode, 
    "Warehouse Name": record.ttbh, 
    "Số RO": record.soRO, 
    "BH/DV": record.bhDv, 
    "Mã LK": record.maLK, 
    "Product Name": record.tenLK, 
    "Model": record.model,
    "Type": record.phanLoai, 
    "Slg": record.slg, 
    "Remark": record.remark
  }));
  const wsScanned = window.XLSX.utils.json_to_sheet(exportScannedData);
  wsScanned['!cols'] = [{ wch: 25 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
  window.XLSX.utils.book_append_sheet(workbook, wsScanned, "LichSu_DaScan");

  const unscannedData = detailedProgressList.filter(r => !r.isScanned).map(record => ({
    "Trạng thái": "Chưa Scan", 
    "Cột SP": record["SP"], 
    "SC Code": record["SC Code"] || record["SC code"] || '', 
    "Warehouse Name": record["Warehouse Name"], 
    "Số RO": record["After-sales work order No."], 
    "BH/DV": record["Repair Type"], 
    "Mã LK": record["Defective material code"], 
    "Product Name": record["Product Name"], 
    "Model": record["Product Model"] || record["Model"] || '', 
    "Type": record["Type"], 
    "Slg": record["Consumed quantity"] || record["Consumed"] || '', 
    "Remark": getRemark(record)
  }));
  const wsUnscanned = window.XLSX.utils.json_to_sheet(unscannedData);
  wsUnscanned['!cols'] = [{ wch: 15 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
  window.XLSX.utils.book_append_sheet(workbook, wsUnscanned, "DanhSach_ChuaScan");

  if (!workbook.Workbook) workbook.Workbook = {};
  workbook.Workbook.Sheets = [{ Hidden: 1 }, { Hidden: 0 }, { Hidden: 0 }];

  return workbook;
};

// ----------------------------------------------------------------------
// MAIN COMPONENT
// ----------------------------------------------------------------------
export default function App() {
  const [excelData, setExcelData] = useState([]);
  const [warehouses, setWarehouses] = useState([]);
  const [selectedWarehouse, setSelectedWarehouse] = useState('');
  const [scanInput, setScanInput] = useState('');
  const [scannedRecords, setScannedRecords] = useState([]);
  const [isLoadingData, setIsLoadingData] = useState(false);
  
  // Custom Searchable Dropdown States
  const [isWhDropdownOpen, setIsWhDropdownOpen] = useState(false);
  const [whSearchQuery, setWhSearchQuery] = useState('');
  const whDropdownRef = useRef(null);

  // URL Cấu hình Google Apps Script của bạn - Được lưu cứng để bảo mật và tinh gọn giao diện
  const gasUrl = 'https://script.google.com/macros/s/AKfycbxQ-D2MG2uV7ijkRLKnyfkWCn1IkEYl7uy7r2XtS3mF6ysi2triCBUQfxbXkH8HXZNG/exec';
  const [isUploading, setIsUploading] = useState(false);

  // Password Security Modal States
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const passwordInputRef = useRef(null);

  // Modals & UI States
  const [toast, setToast] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [warningAlert, setWarningAlert] = useState(null); 
  const [pendingQRCheck, setPendingQRCheck] = useState(null); 
  const [secondScanInput, setSecondScanInput] = useState('');
  const [showDetailedProgressModal, setShowDetailedProgressModal] = useState(false);
  
  // CÁC STATE CHO TÍNH NĂNG "BẮT BUỘC QUÉT ĐỦ MÃ LK"
  const [lockedMaLK, setLockedMaLK] = useState(null);
  const [violationCount, setViolationCount] = useState(0);
  const [showForceRemarkModal, setShowForceRemarkModal] = useState(false);
  const [missingRowsForLockedLK, setMissingRowsForLockedLK] = useState([]);
  const [rowRemarks, setRowRemarks] = useState({}); 
  const [selectedMissingRows, setSelectedMissingRows] = useState({}); 

  // State cho việc báo mất xác nhanh (Không xác LK) trực tiếp từ danh sách chưa scan
  const [quickNoXacRow, setQuickNoXacRow] = useState(null);
  const [quickNoXacReason, setQuickNoXacReason] = useState('');

  const mainInputRef = useRef(null);
  const secondInputRef = useRef(null);

  const hasErrorLock = useMemo(() => {
    return scannedRecords.some(r => r.status !== "Khớp, Trả Xác LK về" && r.status !== "Không xác LK");
  }, [scannedRecords]);

  const currentDisplayedRecords = useMemo(() => {
    return scannedRecords.filter(r => r.ttbh === selectedWarehouse || r.ttbh === '');
  }, [scannedRecords, selectedWarehouse]);

  // Lấy SC Code làm tên Folder
  const currentScCode = useMemo(() => {
    const match = excelData.find(row => row["Warehouse Name"] === selectedWarehouse);
    return match ? String(match["SC Code"] || match["SC code"] || "").trim() : '';
  }, [excelData, selectedWarehouse]);

  const filteredWarehouses = useMemo(() => {
    return warehouses.filter(wh => 
      wh.toLowerCase().includes(whSearchQuery.toLowerCase())
    );
  }, [warehouses, whSearchQuery]);

  const detailedProgressList = useMemo(() => {
    if (!selectedWarehouse || excelData.length === 0) return [];
    
    const scannedSPs = scannedRecords
      .filter(r => (r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK") && r.ttbh === selectedWarehouse)
      .map(r => r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
      
    const result = excelData
      .filter(row => row["Warehouse Name"] === selectedWarehouse)
      .map(row => {
        const isScanned = scannedSPs.includes(String(row["SP"]).trim());
        return { ...row, isScanned };
      });
      
    return result.sort((a, b) => (a.isScanned === b.isScanned ? 0 : a.isScanned ? 1 : -1));
  }, [excelData, selectedWarehouse, scannedRecords]);

  // Kiểm tra xem trong các dòng thiếu có dòng nào được phép "Không Xác" không?
  const canSkipAnyMissing = useMemo(() => {
    return missingRowsForLockedLK.some(row => {
        const repairType = String(row["Repair Type"]).trim().toUpperCase();
        const typeLK = String(row["Type"]).trim().toUpperCase();
        return !(repairType === 'IW' && typeLK !== 'PIN');
    });
  }, [missingRowsForLockedLK]);

  // Theo dõi mã LK chưa hoàn thành (Đang dở dang)
  useEffect(() => {
    if (!selectedWarehouse || excelData.length === 0 || scannedRecords.length === 0) {
      setLockedMaLK(null);
      setViolationCount(0);
      return;
    }

    const lastValidRecord = scannedRecords.find(r => 
      r.ttbh === selectedWarehouse && 
      (r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK")
    );

    if (lastValidRecord) {
      const currentMaLK = lastValidRecord.maLK;
      const repairType = lastValidRecord.bhDv;

      if (repairType === 'IW' || repairType === 'OOW') {
         const totalRowsForMaLK = excelData.filter(row => 
           row["Warehouse Name"] === selectedWarehouse && 
           row["Defective material code"] === currentMaLK &&
           String(row["Repair Type"]).trim().toUpperCase() === repairType
         );

         const scannedRowsForMaLK = scannedRecords.filter(r => 
           r.ttbh === selectedWarehouse && 
           r.maLK === currentMaLK && 
           String(r.bhDv).trim().toUpperCase() === repairType &&
           (r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK")
         );

         if (scannedRowsForMaLK.length < totalRowsForMaLK.length && totalRowsForMaLK.length > 0) {
           setLockedMaLK(currentMaLK);
           const scannedSPs = scannedRowsForMaLK.map(r => r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
           const missing = totalRowsForMaLK.filter(row => !scannedSPs.includes(String(row["SP"]).trim()));
           setMissingRowsForLockedLK(missing);
         } else {
           setLockedMaLK(null);
           setViolationCount(0);
         }
      } else {
         setLockedMaLK(null);
         setViolationCount(0);
      }
    } else {
      setLockedMaLK(null);
      setViolationCount(0);
    }
  }, [scannedRecords, selectedWarehouse, excelData]);

  const stats = useMemo(() => {
    let totals = { IW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, OOW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, KMH: { LCD: 0, MAIN: 0, OTHERS: 0 } };
    let scanned = { IW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, OOW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, KMH: { LCD: 0, MAIN: 0, OTHERS: 0 } };

    if (!selectedWarehouse) return { totals, scanned };

    const processItem = (row, targetObj, qtyStr, repairStr, typeStr, kmhStr) => {
      const qty = parseInt(qtyStr) || 1;
      const rType = repairStr === 'IW' ? 'IW' : (repairStr === 'OOW' ? 'OOW' : null);
      const tType = (typeStr === 'LCD' || typeStr === 'MAIN') ? typeStr : 'OTHERS';
      
      if (rType && targetObj[rType]) {
        targetObj[rType][tType] += qty;
        targetObj[rType].total += qty;
      }
      if (kmhStr === "1" || kmhStr.includes("KMH")) {
        targetObj.KMH[tType] += qty;
      }
    };

    excelData.forEach(row => {
      if (row["Warehouse Name"] === selectedWarehouse) {
        processItem(row, totals, row["Consumed quantity"] || row["Consumed"], String(row["Repair Type"]).trim().toUpperCase(), String(row["Type"]).trim().toUpperCase(), String(row["KMH"]));
      }
    });

    currentDisplayedRecords.forEach(record => {
      if (record.status === "Khớp, Trả Xác LK về" || record.status === "Không xác LK") {
        processItem(record, scanned, record.slg, String(record.bhDv).trim().toUpperCase(), String(record.phanLoai).trim().toUpperCase(), record.remark);
      }
    });

    return { totals, scanned };
  }, [excelData, selectedWarehouse, currentDisplayedRecords]);

  const showToast = useCallback((message, type = 'info') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  }, []);

  const fetchGoogleSheetData = useCallback(async () => {
    setIsLoadingData(true);
    await new Promise(resolve => setTimeout(resolve, 50)); 

    try {
      const sheetId = "10zbQkv7f_7nwVawF5lFDdqfpbC23UT1YBjgJ3aZyNWU";
      const url = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
      const response = await fetch(url);
      
      if (!response.ok) throw new Error("Network response was not ok");
      const data = await response.arrayBuffer();

      await new Promise(resolve => setTimeout(resolve, 50)); 

      if (!window.XLSX) {
        showToast("Thư viện xử lý chưa sẵn sàng, vui lòng tải lại trang.", "warning");
        setIsLoadingData(false);
        return;
      }

      const workbook = window.XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames.includes("DataXacLK") ? "DataXacLK" : workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rawJsonData = window.XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      if (rawJsonData.length === 0) {
        showToast("Dữ liệu trống!", "error");
        setIsLoadingData(false);
        return;
      }

      const jsonData = rawJsonData.map(row => {
        const cleanRow = {};
        Object.keys(row).forEach(key => cleanRow[key.trim()] = row[key]);
        return {
          ...cleanRow,
          "SP": String(cleanRow["SP"] || "").split(' ').join('').toUpperCase(),
          "Warehouse Name": String(cleanRow["Warehouse Name"] || "").trim()
        };
      });

      setExcelData(jsonData);
      
      const whList = [...new Set(jsonData.map(item => item["Warehouse Name"]).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'vi'));
        
      setWarehouses(whList);
      
      if (whList.length > 0) {
        setSelectedWarehouse(prev => prev || whList[0]);
      }
      
      setTimeout(() => { if (mainInputRef.current) mainInputRef.current.focus(); }, 100);
    } catch (error) {
      showToast("Lỗi kết nối tải dữ liệu! Vui lòng thử lại sau.", "error");
    } finally {
      setIsLoadingData(false);
    }
  }, [showToast]);

  useEffect(() => {
    function handleClickOutside(event) {
      if (whDropdownRef.current && !whDropdownRef.current.contains(event.target)) {
        setIsWhDropdownOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // Autofocus cho hộp nhập Password
  useEffect(() => {
    if (showPasswordModal && passwordInputRef.current) {
      passwordInputRef.current.focus();
    }
  }, [showPasswordModal]);

  useEffect(() => {
    if (!window.XLSX) {
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.async = true;
      script.onload = fetchGoogleSheetData;
      document.body.appendChild(script);
    } else {
      fetchGoogleSheetData();
    }
  }, [fetchGoogleSheetData]);

  useEffect(() => {
    const isModalOpen = pendingQRCheck || showDetailedProgressModal || confirmDialog || isWhDropdownOpen || showForceRemarkModal || warningAlert || showPasswordModal || quickNoXacRow || isUploading;
    if (!isModalOpen && !hasErrorLock && mainInputRef.current) {
      mainInputRef.current.focus();
    }
  }, [pendingQRCheck, hasErrorLock, scannedRecords.length, showDetailedProgressModal, confirmDialog, selectedWarehouse, isWhDropdownOpen, showForceRemarkModal, warningAlert, showPasswordModal, quickNoXacRow, isUploading]);

  const handleMainScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    
    const rawScan = String(scanInput).trim();
    if (!rawScan) return;

    if (hasErrorLock) {
      showToast("Hệ thống đang khóa! Vui lòng xóa dòng lỗi để tiếp tục.", "error");
      return;
    }
    if (excelData.length === 0) {
      showToast("Dữ liệu gốc chưa tải xong, vui lòng chờ trong giây lát!", "warning");
      setScanInput('');
      return;
    }
    if (!selectedWarehouse) {
      showToast("Vui lòng chọn TTBH!", "warning");
      setScanInput('');
      return;
    }

    let isQRFirstScan = false;
    let matchedRow = null;
    let extractedSPString = '';

    // LUÔN LUÔN tìm theo format số phiếu (SP) trước để chống lỗi format gạch dưới
    extractedSPString = String(rawScan).split('_').join('').split(' ').join('').toUpperCase();
    matchedRow = excelData.find(row => row["SP"] === extractedSPString && row["Warehouse Name"] === selectedWarehouse);

    // Nếu không tìm thấy, thử tìm theo QR Code nguyên gốc
    if (!matchedRow) {
       matchedRow = excelData.find(row => String(row["QRCode"]).trim() === rawScan && row["Warehouse Name"] === selectedWarehouse && rawScan !== '');
       if (matchedRow) {
           extractedSPString = String(matchedRow["SP"]).trim();
           isQRFirstScan = true;
       }
    }

    if (!matchedRow) {
       setScannedRecords(prev => [createRecord(rawScan, "Không đúng Xác LK hoặc QR ko tồn tại"), ...prev]);
       setScanInput('');
       return;
    }

    if (lockedMaLK && matchedRow["Defective material code"] !== lockedMaLK) {
        const newVCount = violationCount + 1;
        setViolationCount(newVCount);
        
        let msg = `Bạn phải quét cho xong toàn bộ Mã LK: ${lockedMaLK}.\nCòn thiếu ${missingRowsForLockedLK.length} số phiếu chưa scan.\n(Cảnh báo lần ${newVCount})`;
        
        if (newVCount >= 3) {
           setShowForceRemarkModal(true);
           setScanInput('');
           return;
        } else {
           setWarningAlert(msg);
           setScanInput('');
           return;
        }
    }

    const isDuplicate = currentDisplayedRecords.some(r => (r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === extractedSPString);
    if (isDuplicate) {
       setScannedRecords(prev => [createRecord(rawScan, "Bị trùng, SP hoặc QR này đã được scan"), ...prev]);
       setScanInput('');
       return;
    }

    const repairType = String(matchedRow["Repair Type"]).trim().toUpperCase();
    const type = String(matchedRow["Type"]).trim().toUpperCase();

    if (repairType === 'IW' && (type === 'MAIN' || type === 'LCD')) {
      if (!isQRFirstScan) {
         setPendingQRCheck({ type: 'WAITING_FOR_QR', rowData: matchedRow, spScan: rawScan });
      } else {
         setPendingQRCheck({ type: 'WAITING_FOR_SP', rowData: matchedRow, qrScan: rawScan });
      }
      setTimeout(() => { if (secondInputRef.current) secondInputRef.current.focus(); }, 100);
    } else {
      setScannedRecords(prev => [createRecord(rawScan, "Khớp, Trả Xác LK về", matchedRow), ...prev]);
    }
    
    setScanInput('');
  }, [scanInput, hasErrorLock, excelData, selectedWarehouse, currentDisplayedRecords, lockedMaLK, violationCount, missingRowsForLockedLK, showToast]);

  const handleSecondScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    const secondScan = String(secondScanInput).trim();
    if (!secondScan) return;

    if (hasErrorLock) {
       showToast("Hệ thống đang khóa! Xóa lỗi bên dưới để tiếp tục.", "error");
       return;
    }

    const { type, rowData, spScan, qrScan } = pendingQRCheck;

    if (type === 'WAITING_FOR_QR') {
      const expectedQR = String(rowData["QRCode"]).trim();
      const isMatched = secondScan === expectedQR;
      const status = isMatched ? "Khớp, Trả Xác LK về" : "QR Tem Xác và LK ko khớp nhau";
      setScannedRecords(prev => [createRecord(spScan, status, rowData, { secondQRScanned: secondScan, expectedQR }), ...prev]);
    } 
    else if (type === 'WAITING_FOR_SP') {
      const spString = String(secondScan).split('_').join('').split(' ').join('').toUpperCase();
      const expectedSP = String(rowData["SP"]).trim();
      const isMatched = spString === expectedSP;
      const status = isMatched ? "Khớp, Trả Xác LK về" : "Số phiếu/Mã LK không khớp với QR Tem Xác";

      if (isMatched) {
        const isDuplicate = currentDisplayedRecords.some(r => (r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === spString);
        if (isDuplicate) {
           setScannedRecords(prev => [createRecord(secondScan, "Bị trùng, phải xóa dòng này"), ...prev]);
        } else {
           setScannedRecords(prev => [createRecord(secondScan, status, rowData, { qrScan: qrScan }), ...prev]);
        }
      } else {
        setScannedRecords(prev => [createRecord(secondScan, status, rowData), ...prev]);
      }
    }

    setPendingQRCheck(null);
    setSecondScanInput('');
  }, [secondScanInput, pendingQRCheck, currentDisplayedRecords, hasErrorLock, showToast]);

  const handleDeleteRecord = useCallback((id) => {
    setScannedRecords(prev => prev.filter(record => record.id !== id));
  }, []);

  // Xử lý nút Xác nhận "Không xác LK" hàng loạt
  const handleBulkForceRemark = () => {
    const newRecords = [];
    let hasError = false;

    Object.keys(selectedMissingRows).forEach(idx => {
       if (selectedMissingRows[idx]) {
          const remarkText = rowRemarks[idx];
          if (!remarkText || !remarkText.trim()) {
             hasError = true;
          } else {
             const selectedRow = missingRowsForLockedLK[idx];
             const rawSP = String(selectedRow["SP"]);
             const newRecord = createRecord(rawSP, "Không xác LK", selectedRow);
             newRecord.remark = remarkText.trim();
             newRecords.push(newRecord);
          }
       }
    });

    if (hasError) {
       showToast("Vui lòng nhập đầy đủ lý do cho các phiếu đã được tick chọn!", "warning");
       return;
    }
    if (newRecords.length === 0) {
       showToast("Vui lòng tick chọn ít nhất 1 phiếu để xác nhận!", "warning");
       return;
    }

    setScannedRecords(prev => [...newRecords, ...prev]);
    setShowForceRemarkModal(false);
    setViolationCount(0);
    setRowRemarks({});
    setSelectedMissingRows({});
    showToast(`Đã ghi nhận ${newRecords.length} phiếu Không xác LK thành công.`, "success");
  };

  // Xác nhận Không Xác LK nhanh từ danh sách Chưa Scan
  const handleQuickNoXacSubmit = () => {
    if (!quickNoXacReason.trim()) {
      showToast("Vui lòng điền lý do báo mất linh kiện!", "warning");
      return;
    }
    const rawSP = String(quickNoXacRow["SP"]);
    const newRecord = createRecord(rawSP, "Không xác LK", quickNoXacRow);
    newRecord.remark = quickNoXacReason.trim();

    setScannedRecords(prev => [newRecord, ...prev]);
    setQuickNoXacRow(null);
    setQuickNoXacReason('');
    showToast(`Đã ghi nhận Không xác LK cho số RO: ${newRecord.soRO}`, "success");
  };

  const handleSaveSession = useCallback(() => {
    if (scannedRecords.length === 0) {
      showToast("Chưa có dữ liệu scan nào để lưu phiên!", "warning");
      return;
    }
    const dataStr = JSON.stringify(scannedRecords, null, 2);
    const blob = new Blob([dataStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    
    link.href = url;
    link.download = `PhienScan_${selectedWarehouse ? selectedWarehouse.split(' ').join('') : 'All'}_${new Date().toISOString().slice(0, 10).split('-').join('')}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast("Đã tải xuống file sao lưu phiên làm việc.", "success");
  }, [scannedRecords, selectedWarehouse, showToast]);

  const handleLoadSession = useCallback((e) => {
    const inputTarget = e.target;
    const file = inputTarget.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const loadedData = JSON.parse(evt.target.result);
        if (!Array.isArray(loadedData)) throw new Error("Invalid format");

        const loadedWarehouses = [...new Set(loadedData.map(r => String(r.ttbh || "").trim()).filter(Boolean))];
        
        setWarehouses(prev => {
          const combined = [...new Set([...prev, ...loadedWarehouses])];
          return combined.sort((a, b) => a.localeCompare(b, 'vi'));
        });

        if (loadedWarehouses.length === 1 && selectedWarehouse !== loadedWarehouses[0]) {
          setSelectedWarehouse(loadedWarehouses[0]);
        } else if (!selectedWarehouse && loadedWarehouses.length > 0) {
          setSelectedWarehouse(loadedWarehouses[0]);
        }

        setScannedRecords(prev => {
          const existingScans = new Set(prev.map(r => String(r.rawScan)));
          const normalizedLoadedData = loadedData.map(r => ({
             ...r,
             sp: r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase()
          }));
          const newRecords = normalizedLoadedData.filter(r => !existingScans.has(String(r.rawScan)));
          
          if (newRecords.length > 0) showToast(`Đã nạp thành công ${newRecords.length} dòng dữ liệu mới!`, "success");
          else showToast(`Tất cả ${loadedData.length} dòng đã có sẵn, không có dữ liệu mới.`, "warning");

          return [...newRecords, ...prev];
        });
      } catch (err) {
        showToast("Lỗi khi đọc file phiên làm việc!", "error");
      } finally {
        inputTarget.value = null; 
      }
    };
    reader.readAsText(file);
  }, [selectedWarehouse, showToast]);

  // HÀM ĐỒNG BỘ TRỰC TIẾP LÊN GOOGLE DRIVE BẰNG SIMPLE REQUEST (Không dùng mode: cors)
  const executeUploadToDrive = async (workbook) => {
    if (!gasUrl.trim()) return false;
    
    setIsUploading(true);
    try {
      const base64Data = window.XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
      const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
      const filename = `BaoCao_XacLK_${folderName}.xlsx`;

      // Tạo simple request: không truyền mode, không header tùy chỉnh phức tạp
      // JSON.stringify sẽ tự động thiết lập text/plain làm Content-Type để né preflight OPTIONS
      const response = await fetch(gasUrl, {
        method: "POST",
        body: JSON.stringify({
          warehouseCode: folderName,
          filename: filename,
          fileData: base64Data
        })
      });

      const textRes = await response.text();
      
      // Xử lý an toàn phòng trường hợp Google trả về trang HTML cảnh báo/đăng nhập thay vì JSON
      try {
        const res = JSON.parse(textRes);
        if (res.success) {
           return true;
        } else {
           console.error("Lỗi từ Apps Script:", res.error);
           return false;
        }
      } catch (parseError) {
        console.error("Lỗi định dạng trả về (thường do Script chưa được cấp quyền Anyone):", textRes);
        return false;
      }
    } catch (err) {
      console.error("Lỗi fetch API:", err);
      return false;
    } finally {
      setIsUploading(false);
    }
  };

  const handleUploadToDrive = async () => {
    if (currentDisplayedRecords.length === 0 && detailedProgressList.length === 0) {
      showToast("Không có dữ liệu đối chiếu nào để đồng bộ!", "warning");
      return;
    }
    if (!window.XLSX) {
      showToast("Thư viện Excel đang được tải, vui lòng chờ trong giây lát.", "warning");
      return;
    }

    const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
    const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
    
    const success = await executeUploadToDrive(workbook);
    
    if (success) {
      showToast(`Đẩy lên Google Drive thành công vào thư mục [${folderName}]!`, "success");
    } else {
      showToast("Đồng bộ thất bại! Hãy chắc chắn Web App của bạn đã thiết lập quyền truy cập là 'Mọi người' (Anyone).", "error");
    }
  };

  const handleExportExcel = useCallback(async () => {
    if (currentDisplayedRecords.length === 0 && detailedProgressList.length === 0) {
      showToast("Không có dữ liệu để xuất Excel!", "warning");
      return;
    }
    if (!window.XLSX) {
      showToast("Thư viện Excel đang được tải, vui lòng thử lại.", "warning");
      return;
    }
    const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
    const dateStr = new Date().toISOString().slice(0, 10).split('-').join('');
    
    window.XLSX.writeFile(workbook, `BaoCao_ScanLK_${selectedWarehouse ? selectedWarehouse.split(' ').join('') : 'All'}_${dateStr}.xlsx`);
    showToast("Đã tải xuống file báo cáo Excel.", "success");

    if (gasUrl.trim()) {
      const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
      const uploadSuccess = await executeUploadToDrive(workbook);
      if (uploadSuccess) {
        showToast(`Đã đẩy tự động lên GG Drive thành công vào thư mục [${folderName}]!`, "success");
      } else {
        showToast("Tự động đồng bộ lỗi! Hãy kiểm tra quyền 'Mọi người' (Anyone) của Web App.", "warning");
      }
    }
  }, [currentDisplayedRecords, detailedProgressList, selectedWarehouse, currentScCode, showToast]);

  // HÀM XUẤT THỰC TẾ CHO NÚT "TẢI TẤT CẢ" (Chạy sau khi nhập đúng Password)
  const executeExportAllExcel = useCallback(() => {
    const workbook = window.XLSX.utils.book_new();
    const globalScannedSPs = scannedRecords
      .filter(r => r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK")
      .map(r => r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase());

    const globalDetailedList = excelData.map(row => {
      const isScanned = globalScannedSPs.includes(String(row["SP"]).trim());
      return { ...row, isScanned };
    });

    const statsData = [];
    warehouses.forEach(wh => {
      const whData = globalDetailedList.filter(r => r["Warehouse Name"] === wh);
      if (whData.length === 0) return;

      const iwData = whData.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW');
      const iwGroups = {};
      iwData.forEach(r => {
        const t = String(r["Type"]).toUpperCase();
        const type = (t === 'LCD' || t === 'MAIN') ? t : 'OTHERS';
        const key = `${type}|${r["Defective material code"]}|${r["Product Name"]}`;
        if (!iwGroups[key]) iwGroups[key] = { type, maLK: r["Defective material code"], tenLK: r["Product Name"], total: 0, scanned: 0 };
        iwGroups[key].total += 1;
        if (r.isScanned) iwGroups[key].scanned += 1;
      });

      Object.values(iwGroups).forEach(g => {
        statsData.push({ "TTBH": wh, "Phân Loại": "IW", "Nhóm LK": g.type, "Mã LK": g.maLK, "Tên LK": g.tenLK, "Cần Scan": g.total, "Đã Scan": g.scanned, "Còn Thiếu": g.total - g.scanned });
      });

      const oowData = whData.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW');
      const oowGroups = { 'LCD': { total: 0, scanned: 0 }, 'MAIN': { total: 0, scanned: 0 }, 'OTHERS': { total: 0, scanned: 0 } };
      oowData.forEach(r => {
        const t = String(r["Type"]).toUpperCase();
        const type = (t === 'LCD' || t === 'MAIN') ? t : 'OTHERS';
        oowGroups[type].total += 1;
        if (r.isScanned) oowGroups[type].scanned += 1;
      });
      ['LCD', 'MAIN', 'OTHERS'].forEach(type => {
        if (oowGroups[type].total > 0) statsData.push({ "TTBH": wh, "Phân Loại": "OOW", "Nhóm LK": type, "Mã LK": "-", "Tên LK": `Tổng số lượng ${type} (OOW)`, "Cần Scan": oowGroups[type].total, "Đã Scan": oowGroups[type].scanned, "Còn Thiếu": oowGroups[type].total - oowGroups[type].scanned });
      });

      const otherData = whData.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase()));
      const otherGroups = { 'LCD': { total: 0, scanned: 0 }, 'MAIN': { total: 0, scanned: 0 }, 'OTHERS': { total: 0, scanned: 0 } };
      otherData.forEach(r => {
        const t = String(r["Type"]).toUpperCase();
        const type = (t === 'LCD' || t === 'MAIN') ? t : 'OTHERS';
        otherGroups[type].total += 1;
        if (r.isScanned) otherGroups[type].scanned += 1;
      });
      ['LCD', 'MAIN', 'OTHERS'].forEach(type => {
        if (otherGroups[type].total > 0) statsData.push({ "TTBH": wh, "Phân Loại": "Khác", "Nhóm LK": type, "Mã LK": "-", "Tên LK": `Tổng số lượng ${type} (Khác)`, "Cần Scan": otherGroups[type].total, "Đã Scan": otherGroups[type].scanned, "Còn Thiếu": otherGroups[type].total - otherGroups[type].scanned });
      });
    });

    const wsStats = window.XLSX.utils.json_to_sheet(statsData);
    wsStats['!cols'] = [{ wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 45 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
    window.XLSX.utils.book_append_sheet(workbook, wsStats, "ThongKe_TongHop");

    // THỐNG KÊ THEO MODEL (TOÀN QUỐC)
    const modelGroups = {};
    globalDetailedList.forEach(r => {
        const model = String(r["Product Model"] || r["Model"] || 'Unknown').trim();
        const type = String(r["Type"] || 'Unknown').trim();
        const repairType = String(r["Repair Type"] || 'Unknown').trim();
        const slg = parseInt(r["Consumed quantity"]) || parseInt(r["Consumed"]) || 1;

        const key = `${model}|${type}|${repairType}`;
        if (!modelGroups[key]) {
            modelGroups[key] = { model, type, repairType, totalSlg: 0, scannedSlg: 0 };
        }
        modelGroups[key].totalSlg += slg;
        if (r.isScanned) {
            modelGroups[key].scannedSlg += slg;
        }
    });

    const modelStatsData = Object.values(modelGroups).map(g => ({
        "Model": g.model,
        "Type": g.type,
        "BH/DV": g.repairType,
        "Tổng Cần Scan (Slg)": g.totalSlg,
        "Đã Scan Khớp (Slg)": g.scannedSlg,
        "Còn Thiếu (Slg)": g.totalSlg - g.scannedSlg
    })).sort((a, b) => a.Model.localeCompare(b.Model));

    const wsModelStats = window.XLSX.utils.json_to_sheet(modelStatsData);
    wsModelStats['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 20 }];
    window.XLSX.utils.book_append_sheet(workbook, wsModelStats, "ThongKe_TheoModel");

    // ĐỒNG NHẤT CỘT DỮ LIỆU ĐÃ SCAN (ALL)
    const exportScannedData = scannedRecords.map(record => ({
      "Trạng thái": record.status, 
      "Cột SP": record.sp || record.rawScan, 
      "SC Code": record.scCode, 
      "Warehouse Name": record.ttbh, 
      "Số RO": record.soRO, 
      "BH/DV": record.bhDv, 
      "Mã LK": record.maLK, 
      "Product Name": record.tenLK, 
      "Model": record.model,
      "Type": record.phanLoai, 
      "Slg": record.slg, 
      "Remark": record.remark
    }));
    const wsScanned = window.XLSX.utils.json_to_sheet(exportScannedData);
    wsScanned['!cols'] = [{ wch: 25 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
    window.XLSX.utils.book_append_sheet(workbook, wsScanned, "LichSu_DaScan_TatCa");

    // ĐỒNG NHẤT CỘT DỮ LIỆU CHƯA SCAN (ALL)
    const wsUnscannedData = globalDetailedList.filter(r => !r.isScanned).map(record => ({
      "Trạng thái": "Chưa Scan", 
      "Cột SP": record["SP"], 
      "SC Code": record["SC Code"] || record["SC code"] || '', 
      "Warehouse Name": record["Warehouse Name"], 
      "Số RO": record["After-sales work order No."], 
      "BH/DV": record["Repair Type"], 
      "Mã LK": record["Defective material code"], 
      "Product Name": record["Product Name"], 
      "Model": record["Product Model"] || record["Model"] || '', 
      "Type": record["Type"], 
      "Slg": record["Consumed quantity"] || record["Consumed"] || '', 
      "Remark": getRemark(record)
    }));
    const wsUnscanned = window.XLSX.utils.json_to_sheet(wsUnscannedData);
    wsUnscanned['!cols'] = [{ wch: 15 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
    window.XLSX.utils.book_append_sheet(workbook, wsUnscanned, "DanhSach_ChuaScan_TatCa");

    if (!workbook.Workbook) workbook.Workbook = {};
    workbook.Workbook.Sheets = [{ Hidden: 1 }, { Hidden: 1 }, { Hidden: 0 }, { Hidden: 0 }];

    const dateStr = new Date().toISOString().slice(0, 10).split('-').join('');
    window.XLSX.writeFile(workbook, `BaoCao_ScanLK_TongHop_${dateStr}.xlsx`);
    showToast("Đã tải xuống file báo cáo TỔNG HỢP thành công.", "success");
  }, [excelData, scannedRecords, warehouses, showToast]);

  // NÚT TRIGGER MỞ MODAL PASSWORD
  const handleExportAllExcel = useCallback(() => {
    if (excelData.length === 0 && scannedRecords.length === 0) {
      showToast("Không có dữ liệu để xuất Excel!", "warning");
      return;
    }
    if (!window.XLSX) {
      showToast("Thư viện Excel đang được tải, vui lòng thử lại.", "warning");
      return;
    }
    setShowPasswordModal(true);
  }, [excelData, scannedRecords, showToast]);

  const handlePasswordSubmit = () => {
    if (passwordInput === '11221122a') {
      executeExportAllExcel();
    }
    setPasswordInput('');
    setShowPasswordModal(false);
  };

  const renderStatus = useCallback((status) => {
    if (status === "Khớp, Trả Xác LK về") return <span className="inline-flex items-center text-emerald-600 font-bold"><CheckCircle className="w-4 h-4 mr-1.5" /> Khớp, Trả Xác</span>;
    if (status === "Không xác LK") return <span className="inline-flex items-center text-gray-600 font-bold bg-gray-100 border border-gray-300 px-2 py-1 rounded-md"><AlertTriangle className="w-4 h-4 mr-1.5" /> Không xác LK</span>;
    if (status.includes("trùng")) return <span className="inline-flex items-center text-red-600 font-bold bg-red-100 px-2 py-1 rounded-md"><Ban className="w-4 h-4 mr-1.5" /> Lỗi trùng lặp</span>;
    if (status.includes("định dạng") || status.includes("không tồn tại")) return <span className="inline-flex items-center text-red-600 font-bold bg-red-100 px-2 py-1 rounded-md"><AlertTriangle className="w-4 h-4 mr-1.5" /> Sai định dạng mã</span>;
    return <span className="inline-flex items-center text-rose-500 font-bold bg-rose-50 px-2 py-1 rounded-md"><XCircle className="w-4 h-4 mr-1.5" /> Sai linh kiện / QR</span>;
  }, []);

  const renderDetailedTableGroup = useCallback((title, dataList, titleColorClass) => {
    if (dataList.length === 0) return null;
    return (
      <div className="mb-6">
        <h3 className={`font-bold text-sm mb-3 flex items-center uppercase ${titleColorClass}`}>
          <span className="w-2 h-6 bg-current mr-2 rounded-sm opacity-80"></span>
          Nhóm phiếu {title} 
          <span className="ml-2 bg-gray-100 text-gray-700 px-2.5 py-0.5 rounded-full text-xs border border-gray-200">{dataList.length}</span>
        </h3>
        <table className="w-full text-left border-collapse whitespace-nowrap bg-white rounded-lg overflow-hidden shadow-sm border border-gray-200">
          <thead className="bg-gray-100 font-bold text-xs uppercase tracking-wider text-gray-600 border-b border-gray-200">
            <tr>
              <th className="p-3 px-4">Trạng Thái</th>
              <th className="p-3 px-4">Số RO</th>
              <th className="p-3 px-4">Mã LK</th>
              <th className="p-3 px-4">Tên LK</th>
              <th className="p-3 px-4">Model</th>
              <th className="p-3 px-4">Loại</th>
              <th className="p-3 px-4 text-center">Slg</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {dataList.map((row, idx) => {
              const repairType = String(row["Repair Type"] || '').trim().toUpperCase();
              const typeLK = String(row["Type"] || '').trim().toUpperCase();
              const allowQuickNoXac = repairType === 'OOW' || (repairType === 'IW' && typeLK === 'PIN');

              return (
                <tr key={idx} className={`hover:bg-gray-50/80 transition-colors ${row.isScanned ? 'bg-gray-50/50 opacity-60' : ''}`}>
                  <td className="p-2.5 px-4 w-[280px]">
                    <div className="flex items-center gap-2">
                      {row.isScanned ? (
                        <span className="inline-flex items-center text-gray-500 font-bold bg-gray-100 px-2.5 py-1 rounded text-xs border border-gray-200"><CheckCircle className="w-3.5 h-3.5 mr-1" /> Đã Scan</span>
                      ) : (
                        <>
                          <span className="inline-flex items-center text-red-500 font-bold bg-red-50 px-2.5 py-1 rounded text-xs border border-red-100"><AlertTriangle className="w-3.5 h-3.5 mr-1" /> Chưa Scan</span>
                          {allowQuickNoXac && (
                            <button
                              onClick={() => {
                                setQuickNoXacRow(row);
                                setQuickNoXacReason('');
                              }}
                              className="text-[10px] font-bold bg-red-100 text-red-700 hover:bg-red-200 px-2 py-1 rounded border border-red-200 transition-colors shadow-sm shrink-0"
                            >
                              Không xác LK
                            </button>
                          )}
                        </>
                      )}
                    </div>
                  </td>
                  <td className="p-2.5 px-4">
                    <div className="flex items-center gap-2">
                      <span className="font-mono text-[13px] text-gray-700">{row["After-sales work order No."]}</span>
                      {row["After-sales work order No."] && (
                        <a href={`https://gcsm-sg.oppoit.com/order/order-management/after-sales-order/${row["After-sales work order No."]}/detail`} target="_blank" rel="noopener noreferrer" className="inline-flex items-center text-[10px] font-bold bg-blue-50 text-blue-600 hover:bg-blue-100 hover:text-blue-700 px-1.5 py-0.5 rounded border border-blue-200 transition-colors" title="Kiểm tra hệ thống GCSM">
                          <ExternalLink className="w-2.5 h-2.5 mr-1" /> Check GCSM
                        </a>
                      )}
                    </div>
                  </td>
                  <td className="p-2.5 px-4 font-mono text-[13px] text-gray-700">{row["Defective material code"]}</td>
                  <td className="p-2.5 px-4 text-gray-700 max-w-[200px] truncate font-medium" title={row["Product Name"]}>{row["Product Name"]}</td>
                  <td className="p-2.5 px-4 text-gray-600 text-[13px]">{row["Product Model"] || row["Model"] || ''}</td>
                  <td className="p-2.5 px-4 text-gray-500 text-xs">{row["Type"]}</td>
                  <td className="p-2.5 px-4 font-bold text-gray-700 text-center">{row["Consumed quantity"] || row["Consumed"] || '1'}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }, [quickNoXacRow, quickNoXacReason, showToast]);

  return (
    <div className="h-screen flex flex-col bg-[#f8f9fa] text-sm font-sans text-gray-800 overflow-hidden">
      
      {isLoadingData && (
        <div className="fixed top-0 left-0 w-full h-1 bg-blue-100 z-[300]">
          <div className="h-full bg-blue-600 animate-pulse w-full"></div>
        </div>
      )}

      {toast && (
        <div className={`fixed bottom-6 right-6 z-[300] px-6 py-3 rounded-lg shadow-2xl text-white font-bold transition-all duration-300 transform translate-y-0 opacity-100 ${toast.type === 'error' ? 'bg-red-500' : toast.type === 'success' ? 'bg-emerald-500' : toast.type === 'warning' ? 'bg-amber-500' : 'bg-blue-500'}`}>
          {toast.message}
        </div>
      )}

      {/* Warning Alert Modal */}
      {warningAlert && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 max-w-sm w-full border border-red-100 text-center relative z-50">
            <AlertTriangle className="w-12 h-12 text-amber-500 mx-auto mb-4 animate-bounce" />
            <h3 className="font-bold text-lg text-amber-600 mb-2">CẢNH BÁO (Lần {violationCount})</h3>
            <p className="text-gray-700 whitespace-pre-line mb-6 font-medium">{warningAlert}</p>
            <button onClick={() => setWarningAlert(null)} className="w-full py-2.5 bg-amber-500 text-white rounded-lg hover:bg-amber-600 font-bold transition-colors shadow-md">Tôi đã hiểu</button>
          </div>
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-gray-900/40 backdrop-blur-sm flex items-center justify-center z-[150] p-4">
          <div className="bg-white rounded-xl shadow-2xl p-6 max-w-sm w-full border border-gray-100 relative z-50">
            <h3 className="font-bold text-lg mb-2 text-gray-800">{confirmDialog.title}</h3>
            <p className="text-gray-600 mb-6">{confirmDialog.message}</p>
            <div className="flex justify-end gap-3">
              <button onClick={() => setConfirmDialog(null)} className="px-4 py-2 bg-gray-100 rounded-lg hover:bg-gray-200 font-medium transition-colors">Hủy</button>
              <button onClick={() => { confirmDialog.onConfirm(); setConfirmDialog(null); }} className="px-4 py-2 bg-red-500 text-white rounded-lg hover:bg-red-600 font-medium transition-colors">Đồng ý</button>
            </div>
          </div>
        </div>
      )}

      {/* Password Verification Modal (Silent Fail) */}
      {showPasswordModal && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-2xl shadow-xl p-5 max-w-sm w-full border border-gray-200 relative z-50 text-center">
            <Settings className="w-12 h-12 text-blue-600 mx-auto mb-3 animate-spin" style={{ animationDuration: '4s' }} />
            <h3 className="font-bold text-base text-gray-800 mb-1.5">Xác thực xuất dữ liệu</h3>
            <p className="text-xs text-gray-500 mb-4">Vui lòng nhập mật khẩu xác nhận quyền xuất báo cáo TỔNG HỢP toàn quốc.</p>
            
            <input 
              ref={passwordInputRef}
              type="password"
              placeholder="Nhập mật khẩu..."
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') handlePasswordSubmit(); }}
              className="w-full text-center border-2 border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:border-blue-500 bg-gray-50/50 font-mono mb-4"
            />
            
            <div className="flex gap-2.5">
              <button 
                onClick={() => { setShowPasswordModal(false); setPasswordInput(''); }}
                className="flex-1 py-2 text-xs bg-gray-100 text-gray-600 rounded-lg hover:bg-gray-200 font-bold transition-colors"
              >
                Hủy bỏ
              </button>
              <button 
                onClick={handlePasswordSubmit}
                className="flex-1 py-2 text-xs bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-bold transition-colors shadow-md"
              >
                Xác nhận
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Quick NoXac Modal */}
      {quickNoXacRow && (
        <div className="fixed inset-0 bg-gray-900/80 backdrop-blur-md flex items-center justify-center z-[240] p-4">
          <div className="bg-white rounded-2xl shadow-2xl p-6 max-w-md w-full border border-red-100 relative z-[250]">
            <div className="flex items-center text-red-600 mb-4 border-b border-gray-100 pb-3">
              <AlertTriangle className="w-6 h-6 mr-2 animate-pulse" />
              <h2 className="text-lg font-bold text-gray-800">Khai Báo Không Xác Linh Kiện</h2>
            </div>
            
            <div className="bg-red-50 border border-red-200 rounded-xl p-3.5 mb-5 space-y-1.5 text-xs text-red-800">
              <p><strong>Số RO:</strong> {quickNoXacRow["After-sales work order No."]}</p>
              <p><strong>Mã LK:</strong> {quickNoXacRow["Defective material code"]}</p>
              <p><strong>Tên LK:</strong> {quickNoXacRow["Product Name"]}</p>
              <p><strong>Phân Loại:</strong> {quickNoXacRow["Repair Type"]} - {quickNoXacRow["Type"]}</p>
            </div>

            <div className="mb-6">
              <label className="block text-gray-700 text-xs font-extrabold uppercase mb-2">Nhập lý do báo mất linh kiện (Bắt buộc):</label>
              <input
                type="text"
                value={quickNoXacReason}
                onChange={(e) => setQuickNoXacReason(e.target.value)}
                placeholder="Ví dụ: TTBH báo mất, Khách không trả..."
                className="w-full border-2 border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:border-red-400 bg-gray-50/30 font-medium"
                autoFocus
              />
            </div>

            <div className="flex justify-end gap-2.5">
              <button 
                onClick={() => { setQuickNoXacRow(null); setQuickNoXacReason(''); }}
                className="px-4 py-2 text-xs bg-gray-100 text-gray-600 hover:bg-gray-200 rounded-lg font-bold transition-colors"
              >
                Hủy bỏ
              </button>
              <button 
                onClick={handleQuickNoXacSubmit}
                className="px-4 py-2 text-xs bg-red-600 hover:bg-red-700 text-white rounded-lg font-bold transition-colors shadow-md"
              >
                Xác nhận báo mất
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Force Remark Modal Redesigned */}
      {showForceRemarkModal && (
        <div className="fixed inset-0 bg-gray-900/85 backdrop-blur-md flex items-center justify-center z-[220] p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl border border-gray-100 flex flex-col max-h-[90vh] relative z-50">
            <div className="p-5 border-b border-gray-100 shrink-0 bg-rose-50 rounded-t-xl">
               <h3 className="font-bold text-xl text-rose-700 flex items-center"><AlertTriangle className="w-6 h-6 mr-2" /> Cảnh báo: Phải Scan xong toàn bộ 1 mã LK của IW hoặc OOW</h3>
               <p className="text-rose-600 mt-2 text-sm font-medium">Bạn đã cố tình quét mã linh kiện khác khi <strong>Mã LK: {lockedMaLK}</strong> vẫn chưa được quét xong!</p>
               <p className="text-rose-700 mt-1 text-xs italic font-bold">* Lưu ý: Các loại linh kiện Bảo hành (IW) không phải là PIN bắt buộc phải thu hồi xác, KHÔNG được phép tick chọn "Không xác LK".</p>
            </div>
            
            <div className="p-5 flex-1 overflow-y-auto">
               <p className="font-bold text-gray-700 mb-3">Vui lòng kiểm tra và xác nhận tình trạng các số RO còn thiếu dưới đây:</p>
               <div className="border border-gray-200 rounded-lg overflow-hidden">
                 <table className="w-full text-left text-sm whitespace-nowrap">
                   <thead className="bg-gray-50 border-b border-gray-200">
                     <tr>
                       <th className="p-3 px-4 font-semibold text-gray-600 w-12 text-center">Chọn</th>
                       <th className="p-3 px-4 font-semibold text-gray-600">Số RO</th>
                       <th className="p-3 px-4 font-semibold text-gray-600">Tên LK</th>
                       <th className="p-3 px-4 font-semibold text-gray-600">Loại / Nhóm</th>
                       <th className="p-3 px-4 font-semibold text-gray-600">Nhập Lý do (Bắt buộc)</th>
                     </tr>
                   </thead>
                   <tbody className="divide-y divide-gray-100">
                     {missingRowsForLockedLK.map((row, idx) => {
                        const repairType = String(row["Repair Type"]).trim().toUpperCase();
                        const typeLK = String(row["Type"]).trim().toUpperCase();
                        const isStrictIW = repairType === 'IW' && typeLK !== 'PIN';

                        return (
                          <tr key={idx} className={`transition-colors ${isStrictIW ? 'bg-gray-50/50 opacity-70' : 'hover:bg-gray-50'}`}>
                             <td className="p-2.5 px-4 text-center">
                               <input
                                  type="checkbox"
                                  disabled={isStrictIW}
                                  checked={!!selectedMissingRows[idx]}
                                  onChange={(e) => setSelectedMissingRows({...selectedMissingRows, [idx]: e.target.checked})}
                                  className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer disabled:cursor-not-allowed"
                               />
                             </td>
                             <td className="p-2.5 px-4 font-mono font-semibold text-gray-700">{row["After-sales work order No."]}</td>
                             <td className="p-2.5 px-4 text-gray-600 truncate max-w-[200px]" title={row["Product Name"]}>{row["Product Name"]}</td>
                             <td className="p-2.5 px-4 text-xs font-bold text-gray-500">
                                {repairType} - {typeLK}
                             </td>
                             <td className="p-2.5 px-4">
                               {isStrictIW ? (
                                  <span className="text-red-500 text-[11px] font-bold">Bắt buộc có xác</span>
                               ) : (
                                  <input
                                    type="text"
                                    placeholder="Lý do (TT báo mất...)"
                                    disabled={!selectedMissingRows[idx]}
                                    value={rowRemarks[idx] || ''}
                                    onChange={(e) => setRowRemarks({...rowRemarks, [idx]: e.target.value})}
                                    className="border border-gray-300 rounded px-2 py-1.5 text-xs w-full focus:border-red-400 focus:outline-none bg-white disabled:bg-gray-100 disabled:cursor-not-allowed"
                                  />
                               )}
                             </td>
                          </tr>
                        );
                     })}
                   </tbody>
                 </table>
               </div>
            </div>

            <div className="p-4 border-t border-gray-100 shrink-0 bg-gray-50 rounded-b-xl flex justify-between items-center">
              <span className="text-xs text-gray-500 italic">Hệ thống sẽ mở khóa quét mã tiếp theo khi hoàn tất.</span>
              <div className="flex gap-3">
                 <button onClick={() => { setShowForceRemarkModal(false); setViolationCount(0); setSelectedMissingRows({}); setRowRemarks({}); }} className="px-5 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 font-bold transition-colors">Đóng tạm thời</button>
                 {canSkipAnyMissing && (
                    <button onClick={handleBulkForceRemark} className="px-5 py-2 bg-red-500 text-white rounded-lg hover:bg-red-600 font-bold transition-colors shadow-sm">Xác nhận Không Xác LK</button>
                 )}
              </div>
            </div>
          </div>
        </div>
      )}

      {pendingQRCheck && (
        <div className="fixed inset-0 bg-gray-900/40 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-xl p-6 max-w-md w-full border border-gray-100">
            <div className="flex items-center text-amber-600 mb-4 border-b border-gray-100 pb-3">
              <AlertTriangle className="w-6 h-6 mr-2" />
              <h2 className="text-lg font-bold text-gray-800">Xác Thực 2 Bước (SP & QR)</h2>
            </div>
            <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 mb-5">
              <p className="font-bold text-amber-800 text-sm">
                {pendingQRCheck.type === 'WAITING_FOR_QR' 
                  ? `Xác ${pendingQRCheck.rowData["Repair Type"]} ${pendingQRCheck.rowData["Type"]} yêu cầu 100% phải Scan QR Tem Xác LK!`
                  : `Đã nhận mã QR Tem Xác. Vui lòng Scan Số phiếu/Mã LK (SP) để hoàn tất!`}
              </p>
              <p className="text-xs text-gray-600 mt-1 font-mono">Mã Phiếu: <strong className="bg-white px-1 rounded border border-amber-200">{pendingQRCheck.rowData["After-sales work order No."]}</strong></p>
            </div>
            <div className="mb-5">
              <label className="block text-gray-700 text-sm font-bold mb-2">
                {pendingQRCheck.type === 'WAITING_FOR_QR' ? 'Quét mã QR trên tem xác:' : 'Quét mã SP (Số Phiếu_Mã LK):'}
              </label>
              <input ref={secondInputRef} type="text" maxLength={100} value={secondScanInput} onChange={(e) => setSecondScanInput(e.target.value)} onKeyDown={handleSecondScan} 
                placeholder={pendingQRCheck.type === 'WAITING_FOR_QR' ? 'Bắn mã QR vào đây...' : 'Bắn mã SP vào đây...'} 
                className="w-full border-2 border-amber-300 rounded-lg px-4 py-3 text-base font-mono focus:outline-none focus:border-amber-500 shadow-inner bg-amber-50/30" />
            </div>
            <div className="flex justify-end">
              <button onClick={() => { setPendingQRCheck(null); setSecondScanInput(''); }} className="px-4 py-2 text-sm bg-gray-100 text-gray-600 rounded-md hover:bg-gray-200 font-bold transition-colors">Hủy bỏ</button>
            </div>
          </div>
        </div>
      )}

      {showDetailedProgressModal && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[100] p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-5xl max-h-[85vh] flex flex-col border border-gray-100 relative z-50">
            <div className="flex justify-between items-center p-4 border-b border-gray-100 shrink-0 bg-gray-50 rounded-t-xl">
              <h2 className="text-lg font-bold text-gray-800 flex items-center">
                <ClipboardList className="w-5 h-5 mr-2 text-blue-600" /> Danh sách phiếu CẦN SCAN - <span className="text-blue-700 mx-1">{selectedWarehouse}</span> 
                <span className="text-xs bg-gray-200 text-gray-600 px-2 py-0.5 rounded-full ml-2 font-medium">Tổng: {detailedProgressList.length}</span>
              </h2>
              <button onClick={() => setShowDetailedProgressModal(false)} className="p-1.5 text-gray-400 hover:text-red-500 rounded-md bg-white border border-gray-200 hover:bg-red-50 transition-colors"><X className="w-5 h-5" /></button>
            </div>
            <div className="overflow-auto flex-1 p-5">
              {detailedProgressList.length === 0 ? (
                <div className="text-center text-gray-500 py-10 font-medium">TTBH này chưa có dữ liệu yêu cầu đối chiếu.</div>
              ) : (
                <>
                  {renderDetailedTableGroup("Bảo Hành (IW)", detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW'), "text-emerald-700")}
                  {renderDetailedTableGroup("Dịch Vụ (OOW)", detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW'), "text-blue-700")}
                  {renderDetailedTableGroup("Khác", detailedProgressList.filter(r => !['IW', 'OOW'].includes(String(r["Repair Type"]).toUpperCase())), "text-gray-700")}
                </>
              )}
            </div>
          </div>
        </div>
      )}

      <header className="bg-white border-b border-gray-200 px-5 py-3 shadow-sm flex justify-between items-center shrink-0 z-20 mt-[2px]">
        <h1 className="text-lg font-extrabold flex items-center text-gray-800 tracking-tight">
          <FileSpreadsheet className="mr-2 w-5 h-5 text-blue-600" /> Hệ Thống Đối Chiếu Xác Linh Kiện
        </h1>
        <div className="flex items-center text-gray-500 font-medium">
          {isLoadingData ? (
            <span className="bg-blue-50 text-blue-600 px-3 py-1 rounded-full text-xs border border-blue-200 flex items-center font-bold shadow-sm">
              <Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" /> Đang đồng bộ ngầm...
            </span>
          ) : (
            <span className="bg-gray-100 px-3 py-1 rounded-full text-xs border border-gray-200">
              Tổng dữ liệu gốc: <strong className="text-gray-800">{excelData.length}</strong> dòng
            </span>
          )}
        </div>
      </header>

      <main className="flex-1 flex flex-col p-3 gap-3 overflow-hidden w-full max-w-[1600px] mx-auto z-10">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 shrink-0">
          <div className="lg:col-span-8 bg-white p-3 rounded-xl shadow-sm border border-gray-200 flex flex-col relative z-20">
            <div className="flex justify-between items-center mb-2 border-b border-gray-100 pb-2">
              <div className="flex items-center gap-3">
                <h2 className="font-bold text-gray-700 flex items-center text-sm">
                  <BarChart2 className="w-4 h-4 mr-1.5 text-blue-500" /> Thống kê:
                </h2>
                
                {/* Custom Searchable TTBH Combobox */}
                <div className="relative" ref={whDropdownRef}>
                  <button
                    type="button"
                    onClick={() => setIsWhDropdownOpen(!isWhDropdownOpen)}
                    className="border border-blue-200 rounded-lg px-3 py-1.5 text-xs font-bold bg-blue-50 text-blue-800 focus:outline-none focus:ring-2 focus:ring-blue-400 min-w-[240px] shadow-sm flex items-center justify-between cursor-pointer"
                    disabled={warehouses.length === 0}
                  >
                    <span className="truncate">{selectedWarehouse || "-- Chọn TTBH --"}</span>
                    <Search className="w-3.5 h-3.5 ml-2 text-blue-600 opacity-70 shrink-0" />
                  </button>
                  
                  {isWhDropdownOpen && (
                    <div className="absolute left-0 mt-1 w-full bg-white border border-gray-200 rounded-lg shadow-xl z-50 p-2 max-h-60 flex flex-col min-w-[260px]">
                      <div className="relative mb-1.5 flex items-center">
                        <Search className="absolute left-2.5 w-3.5 h-3.5 text-gray-400 pointer-events-none" />
                        <input
                          type="text"
                          placeholder="Gõ tìm nhanh..."
                          value={whSearchQuery}
                          onChange={(e) => setWhSearchQuery(e.target.value)}
                          className="w-full text-xs border border-gray-200 rounded-md pl-8 pr-2.5 py-1.5 focus:outline-none focus:border-blue-400 focus:ring-1 focus:ring-blue-400 font-medium"
                          autoFocus
                        />
                        {whSearchQuery && (
                          <button 
                            type="button" 
                            onClick={() => setWhSearchQuery('')}
                            className="absolute right-2 text-gray-400 hover:text-gray-600"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                      <div className="overflow-y-auto flex-1 divide-y divide-gray-50">
                        {filteredWarehouses.length === 0 ? (
                          <div className="p-2.5 text-center text-xs text-gray-400 italic">Không tìm thấy TTBH nào</div>
                        ) : (
                          filteredWarehouses.map((wh, idx) => (
                            <button
                              key={idx}
                              type="button"
                              onClick={() => {
                                setSelectedWarehouse(wh);
                                setIsWhDropdownOpen(false);
                                setWhSearchQuery('');
                              }}
                              className={`w-full text-left px-2.5 py-2 text-xs font-semibold transition-colors flex items-center justify-between ${
                                wh === selectedWarehouse 
                                  ? 'bg-blue-50 text-blue-700' 
                                  : 'hover:bg-gray-50 text-gray-700'
                              }`}
                            >
                              <span className="truncate">{wh}</span>
                              {wh === selectedWarehouse && <CheckCircle className="w-3.5 h-3.5 text-blue-600 shrink-0 ml-1.5" />}
                            </button>
                          ))
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {selectedWarehouse && (
                <button onClick={() => setShowDetailedProgressModal(true)} className="text-[11px] font-semibold bg-gray-50 text-gray-700 px-3 py-1.5 rounded-lg border border-gray-200 hover:bg-blue-50 hover:text-blue-700 hover:border-blue-200 transition-colors flex items-center shadow-sm">
                  <ClipboardList className="w-3.5 h-3.5 mr-1.5" /> DS Chưa Scan
                </button>
              )}
            </div>

            {!selectedWarehouse ? (
              <div className="text-gray-400 text-center py-5 text-xs italic bg-gray-50 rounded-lg border border-dashed border-gray-200">
                {isLoadingData ? "Đang đồng bộ dữ liệu ngầm, vui lòng chờ trong giây lát..." : "Vui lòng chọn TTBH ở ô phía trên để xem thống kê"}
              </div>
            ) : (
              <div className="flex flex-wrap items-center justify-between gap-4 mt-1">
                <div className="flex-1 min-w-[200px] max-w-lg">
                  <table className="w-full border-collapse text-center text-xs rounded-lg overflow-hidden border border-gray-200">
                    <thead>
                      <tr>
                        <th className="bg-gray-50 p-1.5 text-gray-600 font-bold border-b border-gray-200 w-1/4">Loại</th>
                        <th className="bg-emerald-50 p-1.5 text-emerald-700 font-bold border-b border-emerald-100 w-1/4">IW</th>
                        <th className="bg-blue-50 p-1.5 text-blue-700 font-bold border-b border-blue-100 w-1/4">OOW</th>
                        <th className="bg-rose-50 p-1.5 text-rose-600 font-bold border-b border-rose-100 w-1/4">"KMH"</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 bg-white">
                      {stats.scanned && ['LCD', 'MAIN', 'OTHERS'].map(type => (
                        <tr key={type} className="hover:bg-gray-50/50">
                          <td className="p-1.5 font-semibold text-gray-700 border-r border-gray-100">{type === 'OTHERS' ? 'Others' : type === 'MAIN' ? 'Main' : type}</td>
                          <td className="p-1.5 font-medium text-emerald-600 border-r border-gray-100">{stats.scanned.IW[type]} / {stats.totals.IW[type]}</td>
                          <td className="p-1.5 font-medium text-blue-600 border-r border-gray-100">{stats.scanned.OOW[type]} / {stats.totals.OOW[type]}</td>
                          <td className="p-1.5 font-medium text-rose-500">{stats.totals.KMH[type] > 0 ? `${stats.scanned.KMH[type]} / ${stats.totals.KMH[type]}` : '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                
                {/* Horizontal Progress Bars */}
                <div className="flex-1 min-w-[180px] px-4 border-l border-gray-100 flex flex-col justify-center shrink-0">
                  {renderProgressBar(stats.scanned.IW.total, stats.totals.IW.total, 'Tiến độ IW', 'IW')}
                  {renderProgressBar(stats.scanned.OOW.total, stats.totals.OOW.total, 'Tiến độ OOW', 'OOW')}
                </div>
              </div>
            )}
          </div>

          <div className="lg:col-span-4 bg-white p-3 rounded-xl shadow-sm border border-gray-200 flex flex-col gap-3 justify-center z-10">
            {/* Quản lý Phiên */}
            <div className="flex-1 flex flex-col justify-center bg-gray-50/50 p-2.5 rounded-lg border border-gray-100">
              <h2 className="font-bold text-gray-700 mb-2 flex items-center text-[11px] uppercase tracking-wider">
                <Save className="w-3.5 h-3.5 mr-1.5 text-indigo-500" /> Quản lý Phiên
              </h2>
              <div className="grid grid-cols-2 gap-2">
                <button onClick={handleSaveSession} className="text-[11px] font-semibold bg-white border border-indigo-200 text-indigo-700 py-1.5 rounded hover:bg-indigo-50 transition-colors flex items-center justify-center shadow-sm">
                  <Save className="w-3.5 h-3.5 mr-1.5" /> Lưu Phiên
                </button>
                <label className="text-[11px] font-semibold bg-white border border-amber-200 text-amber-700 py-1.5 rounded hover:bg-amber-50 transition-colors flex items-center justify-center shadow-sm cursor-pointer mb-0">
                  <UploadCloud className="w-3.5 h-3.5 mr-1.5" /> Nạp Phiên
                  <input type="file" accept=".json" onChange={handleLoadSession} className="hidden" />
                </label>
              </div>
            </div>

            {/* Xuất Dữ Liệu */}
            <div className="flex-1 flex flex-col justify-center bg-gray-50/50 p-2.5 rounded-lg border border-gray-100">
              <h2 className="font-bold text-gray-700 mb-2 flex items-center text-[11px] uppercase tracking-wider">
                <Download className="w-3.5 h-3.5 mr-1.5 text-emerald-500" /> Xuất Dữ Liệu
              </h2>
              <div className="grid grid-cols-2 gap-2 mb-2">
                <button onClick={handleExportExcel} className="text-[11px] font-semibold bg-white border border-emerald-200 text-emerald-600 py-1.5 rounded hover:bg-emerald-50 transition-colors flex items-center justify-center shadow-sm" title="Tải báo cáo của TTBH đang chọn và tự động đồng bộ Drive">
                  <Download className="w-3.5 h-3.5 mr-1.5" /> Tải TTBH
                </button>
                <button onClick={handleExportAllExcel} className="text-[11px] font-semibold bg-white border border-blue-200 text-blue-600 py-1.5 rounded hover:bg-blue-50 transition-colors flex items-center justify-center shadow-sm" title="Tải báo cáo của tất cả TTBH">
                  <Download className="w-3.5 h-3.5 mr-1.5" /> Tải TẤT CẢ
                </button>
              </div>
              
              {/* NÚT ĐỒNG BỘ DRIVE TRỰC TIẾP THỦ CÔNG */}
              <button 
                onClick={handleUploadToDrive} 
                disabled={isUploading}
                className="text-[11px] font-bold bg-emerald-600 text-white border border-emerald-700 py-2 rounded hover:bg-emerald-700 disabled:bg-gray-300 disabled:border-gray-400 disabled:cursor-not-allowed transition-colors flex items-center justify-center shadow"
              >
                {isUploading ? (
                  <>
                    <Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" /> Đang đẩy lên Drive...
                  </>
                ) : (
                  <>
                    <UploadCloud className="w-3.5 h-3.5 mr-1.5" /> Đẩy báo cáo lên Drive
                  </>
                )}
              </button>
            </div>
          </div>
        </div>

        <div className={`bg-white rounded-xl shadow-sm border flex flex-col flex-1 min-h-0 overflow-hidden transition-all duration-300 ${hasErrorLock ? 'border-red-400 ring-2 ring-red-100' : 'border-gray-200'} z-0`}>
          <div className={`p-3 border-b flex gap-4 items-center shrink-0 ${hasErrorLock ? 'bg-red-50' : 'bg-gray-50'}`}>
            <div className="flex-none font-bold text-gray-800 text-sm flex items-center">
              Lịch sử đối chiếu
              <span className="bg-white border border-gray-200 text-gray-600 px-2 py-0.5 rounded-full text-xs ml-2 font-mono shadow-sm">{currentDisplayedRecords.length}</span>
            </div>
            <div className="flex-1 relative flex items-center max-w-lg ml-4">
               {hasErrorLock ? (
                <span className="text-red-600 font-bold text-xs bg-red-100 px-2 py-0.5 rounded flex items-center shadow-sm">
                  <AlertTriangle className="w-3 h-3 mr-1 animate-bounce" /> HỆ THỐNG ĐANG KHÓA
                </span>
               ) : (
                  <span className="absolute left-3 text-gray-400 z-10"><ScanLine className="w-4 h-4" /></span>
               )}
               <input 
                 ref={mainInputRef}
                 type="text" 
                 maxLength={100}
                 value={scanInput}
                 onChange={(e) => setScanInput(e.target.value)}
                 onKeyDown={handleMainScan}
                 placeholder={isLoadingData ? "Đang tải dữ liệu ngầm..." : hasErrorLock ? "Xóa dòng đỏ lỗi bên dưới..." : "Bắn mã vạch vào đây (Tối đa 100 ký tự)..."}
                 disabled={!selectedWarehouse || hasErrorLock || isLoadingData}
                 className={`w-full rounded-lg pl-9 pr-4 py-1.5 text-sm font-mono shadow-inner transition-all focus:outline-none border-2 ${
                   hasErrorLock ? 'border-red-400 bg-red-100 text-red-700 cursor-not-allowed placeholder-red-500/70 font-bold' : isLoadingData ? 'bg-gray-100 border-gray-200 cursor-not-allowed text-gray-400' : 'border-blue-300 bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-100 text-gray-800'
                 }`}
               />
            </div>
            <div className="flex-none ml-auto">
              <button 
                onClick={() => setConfirmDialog({
                  title: "Xác nhận xóa Lịch sử",
                  message: "Bạn có chắc chắn muốn xóa toàn bộ lịch sử scan của TTBH này không?",
                  onConfirm: () => { setScannedRecords(prev => prev.filter(r => r.ttbh !== selectedWarehouse && r.ttbh !== '')); showToast("Đã làm mới danh sách thành công.", "success"); }
                })} 
                className="text-[11px] font-semibold bg-white border border-gray-200 text-gray-500 px-3 py-1.5 rounded-lg hover:bg-red-50 hover:text-red-600 hover:border-red-200 transition-colors shadow-sm flex items-center"
              >
                <Trash2 className="w-3.5 h-3.5 mr-1.5" /> Làm mới danh sách
              </button>
            </div>
          </div>
          
          <div className="flex-1 overflow-y-auto">
            <table className="w-full text-left border-collapse whitespace-nowrap relative">
              <thead className="bg-white sticky top-0 z-10 shadow-sm border-b border-gray-200">
                <tr>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase w-[160px]">Cột Scan QR</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase w-[180px]">Trạng Thái</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">Số RO</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">Mã LK</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">Tên LK</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">Model</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">Loại</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">BH/DV</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase text-center">Slg</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase">Remark</th>
                  <th className="p-2.5 px-4 border-b border-gray-200 font-semibold text-gray-600 text-[11px] uppercase text-center bg-gray-50">Xóa</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100/50">
                {currentDisplayedRecords.length === 0 ? (
                  <tr>
                    <td colSpan="11" className="p-12 text-center text-gray-400 bg-gray-50/50">
                      <ScanLine className="w-8 h-8 mx-auto mb-3 opacity-20" /> 
                      <span className="font-medium text-sm">Chưa có dữ liệu scan nào ở TTBH này.</span>
                    </td>
                  </tr>
                ) : (
                  currentDisplayedRecords.map((record) => {
                    const isErrorRow = record.status !== "Khớp, Trả Xác LK về" && record.status !== "Không xác LK";
                    const repairType = (record.bhDv || '').toUpperCase();
                    let rowColorClass = "hover:bg-gray-50 transition-colors";
                    if (isErrorRow) rowColorClass = "bg-red-50/80 hover:bg-red-100/80";
                    else if (record.status === "Không xác LK") rowColorClass = "bg-gray-50/50";
                    else if (repairType === 'IW') rowColorClass = "bg-emerald-50/40 hover:bg-emerald-100/50";
                    else if (repairType === 'OOW') rowColorClass = "bg-blue-50/40 hover:bg-blue-100/50";

                    return (
                      <tr key={record.id} className={rowColorClass}>
                        <td className="p-2.5 px-4 font-mono text-[13px] text-gray-700">
                          <div className="w-[160px] overflow-hidden text-ellipsis whitespace-nowrap" title={record.rawScan}>{record.rawScan}</div>
                        </td>
                        <td className="p-2.5 px-4">{renderStatus(record.status)}</td>
                        <td className="p-2.5 px-4 font-mono text-[13px]">
                          {record.soRO ? (
                            <a href={`https://gcsm-sg.oppoit.com/order/order-management/after-sales-order/${record.soRO}/detail`} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:text-blue-800 hover:underline transition-colors font-semibold" title="Kiểm tra trên GCSM">{record.soRO}</a>
                          ) : (
                            <span className="text-gray-600">-</span>
                          )}
                        </td>
                        <td className="p-2.5 px-4 font-mono text-[13px] text-gray-600">{record.maLK}</td>
                        <td className="p-2.5 px-4 text-gray-700 max-w-[200px] truncate font-medium" title={record.tenLK}>{record.tenLK}</td>
                        <td className="p-2.5 px-4 text-gray-600 text-[13px]">{record.model}</td>
                        <td className="p-2.5 px-4 text-gray-600 text-xs">{record.phanLoai}</td>
                        <td className="p-2.5 px-4 font-bold text-xs opacity-70">{record.bhDv}</td>
                        <td className="p-2.5 px-4 font-bold text-[13px] text-gray-700 text-center">{record.slg}</td>
                        <td className="p-2.5 px-4 font-bold text-orange-500 text-xs">{record.remark}</td>
                        <td className={`p-1 text-center border-l border-gray-100/50 ${isErrorRow ? 'bg-red-100/50' : ''}`}>
                          <button onClick={() => handleDeleteRecord(record.id)} className={`p-1.5 rounded transition-all ${isErrorRow ? 'text-white bg-red-500 hover:bg-red-600 shadow-sm animate-bounce' : 'text-gray-400 hover:text-red-500 hover:bg-red-50'}`} title="Xóa dòng này"><Trash2 className="w-4 h-4 mx-auto" /></button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  );
}