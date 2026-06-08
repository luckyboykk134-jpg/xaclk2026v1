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
    rawScan: String(rawScan || ''),
    status: String(status || ''),
    sp: rowData ? String(rowData["SP"] || "").trim() : '', 
    scCode: rowData ? String(rowData["SC Code"] || rowData["SC code"] || "").trim() : '', 
    ttbh: rowData ? String(rowData["Warehouse Name"] || "").trim() : '',
    soRO: rowData ? String(rowData["After-sales work order No."] || "").trim() : '',
    maLK: rowData ? String(rowData["Defective material code"] || "").trim() : '',
    tenLK: rowData ? String(rowData["Product Name"] || "").trim() : '',
    model: rowData ? String(rowData["Product Model"] || rowData["Model"] || "").trim() : '', 
    phanLoai: rowData ? String(rowData["Type"] || "").trim() : '',
    bhDv: rowData ? String(rowData["Repair Type"] || "").trim() : '',
    slg: rowData ? String(rowData["Consumed quantity"] || rowData["Consumed"] || "").trim() : '', 
    remark: String(getRemark(rowData) || ""),
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
    <div className="w-full mb-3 last:mb-0">
      <div className="flex justify-between items-end mb-1">
        <span className={`font-bold text-sm uppercase tracking-wider ${colorText}`}>{String(label)}</span>
        <span className={`font-bold text-sm ${colorText}`}>
          {percent}% <span className="text-xs opacity-70 font-medium ml-0.5">({Number(scannedCount)}/{Number(totalCount)})</span>
        </span>
      </div>
      <div className={`w-full h-3 rounded-full ${colorBg} overflow-hidden shadow-inner`}>
        <div className={`h-full rounded-full ${colorFg} transition-all duration-700 ease-out`} style={{ width: `${percent}%` }}></div>
      </div>
    </div>
  );
};

// Hàm tạo cấu trúc file Excel thống nhất giữa các chức năng xuất
const generateWorkbookData = (selectedWarehouse, currentDisplayedRecords, detailedProgressList) => {
  const workbook = window.XLSX.utils.book_new();

  // ĐỒNG BỘ: Kiểm tra các phiếu đã được quét và ép kiểu thành KMH
  const scannedKMH_SPs = new Set();
  currentDisplayedRecords.forEach(record => {
      if ((record.status === "Khớp, Trả Xác LK về" || record.status === "Không xác LK") && String(record.remark).includes("KMH")) {
          scannedKMH_SPs.add(String(record.sp || '').trim().toUpperCase());
      }
  });

  const isKMH = (row) => {
      const sp = String(row["SP"] || '').trim().toUpperCase();
      return String(row["KMH"]) === "1" || String(row["KMH"]).includes("KMH") || scannedKMH_SPs.has(sp);
  };

  const statsData = [
    {
      "Nhóm Phân Loại": "Bảo Hành (IW)",
      "Tổng Cần Scan": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW').length,
      "Đã Scan Khớp": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW' && r.isScanned).length,
      "Còn Thiếu": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW' && !r.isScanned).length,
    },
    {
      "Nhóm Phân Loại": "Dịch Vụ (OOW)",
      "Tổng Cần Scan": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !isKMH(r)).length,
      "Đã Scan Khớp": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && r.isScanned && !isKMH(r)).length,
      "Còn Thiếu": detailedProgressList.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !r.isScanned && !isKMH(r)).length,
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

  // URL Cấu hình Google Apps Script mã hóa cứng
  const gasUrl = 'https://script.google.com/macros/s/AKfycbycwpfjj7jDSIh1GPQH4RFye5mQltWTs-8I5GgFBMOGV61KkbPOaXuC6UJghkLLa_f2/exec';
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
  
  // State quản lý việc nhập pass Vượt Cấp
  const [showOverrideInput, setShowOverrideInput] = useState(false);
  const [overridePassword, setOverridePassword] = useState('');
  // CỜ VƯỢT CẤP TOÀN CỤC: Khi bật (nhập đúng pass), bỏ qua quy tắc phải quét hết mã LK
  const [isGlobalOverrideActive, setIsGlobalOverrideActive] = useState(false);

  // State cho việc báo mất xác nhanh (Không xác LK) trực tiếp từ danh sách chưa scan
  const [quickNoXacRow, setQuickNoXacRow] = useState(null);
  const [quickNoXacReason, setQuickNoXacReason] = useState('');

  // --- THÊM STATE CHO VƯỢT CẤP OOW (KMH) TRONG MODAL ---
  const [showOOWOverrideInput, setShowOOWOverrideInput] = useState(false);
  const [oowOverridePassword, setOowOverridePassword] = useState('');
  const [isOOWOverrideActive, setIsOOWOverrideActive] = useState(false);
  const [oowScanInput, setOowScanInput] = useState('');

  const mainInputRef = useRef(null);
  const secondInputRef = useRef(null);
  const oowScanInputRef = useRef(null);

  const hasErrorLock = useMemo(() => {
    return scannedRecords.some(r => r.status !== "Khớp, Trả Xác LK về" && r.status !== "Không xác LK");
  }, [scannedRecords]);

  const currentDisplayedRecords = useMemo(() => {
    return scannedRecords.filter(r => r.ttbh === selectedWarehouse || r.ttbh === '');
  }, [scannedRecords, selectedWarehouse]);

  // ----------------------------------------------------------------------
  // PHÁT HIỆN XUNG ĐỘT PHIÊN LÀM VIỆC (QUÉT KHỚP VS KHÔNG XÁC LK TRÙNG SP)
  // ----------------------------------------------------------------------
  const conflictRecordsSPs = useMemo(() => {
    const spMap = new Map(); 
    currentDisplayedRecords.forEach(r => {
      const key = String(r.sp || "").trim().toUpperCase();
      if (!key) return;
      if (!spMap.has(key)) {
        spMap.set(key, { hasKhop: false, hasNoXac: false });
      }
      const info = spMap.get(key);
      if (r.status === "Khớp, Trả Xác LK về") info.hasKhop = true;
      if (r.status === "Không xác LK") info.hasNoXac = true;
    });

    const conflicting = new Set();
    spMap.forEach((val, key) => {
      if (val.hasKhop && val.hasNoXac) {
        conflicting.add(key);
      }
    });
    return conflicting;
  }, [currentDisplayedRecords]);

  const hasConflictLock = useMemo(() => {
    return conflictRecordsSPs.size > 0;
  }, [conflictRecordsSPs]);

  // Sắp xếp lại lịch sử hiển thị: Đẩy các dòng Không xác LK bị xung đột lên trên cùng
  const sortedDisplayedRecords = useMemo(() => {
    return [...currentDisplayedRecords].sort((a, b) => {
      const aSp = String(a.sp || "").trim().toUpperCase();
      const bSp = String(b.sp || "").trim().toUpperCase();
      const aIsConflictNoXac = conflictRecordsSPs.has(aSp) && a.status === "Không xác LK";
      const bIsConflictNoXac = conflictRecordsSPs.has(bSp) && b.status === "Không xác LK";
      
      if (aIsConflictNoXac && !bIsConflictNoXac) return -1;
      if (!aIsConflictNoXac && bIsConflictNoXac) return 1;
      return 0; // Giữ nguyên thứ tự thời gian quét của các dòng khác
    });
  }, [currentDisplayedRecords, conflictRecordsSPs]);

  // Lấy SC Code làm tên Folder
  const currentScCode = useMemo(() => {
    const match = excelData.find(row => String(row["Warehouse Name"]).trim() === selectedWarehouse);
    return match ? String(match["SC Code"] || match["SC code"] || "").trim() : '';
  }, [excelData, selectedWarehouse]);

  const filteredWarehouses = useMemo(() => {
    return warehouses.filter(wh => 
      String(wh).toLowerCase().includes(whSearchQuery.toLowerCase())
    );
  }, [warehouses, whSearchQuery]);

  const detailedProgressList = useMemo(() => {
    if (!selectedWarehouse || excelData.length === 0) return [];
    
    const scannedSPs = scannedRecords
      .filter(r => (r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK") && r.ttbh === selectedWarehouse)
      .map(r => r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
      
    const result = excelData
      .filter(row => String(row["Warehouse Name"]).trim() === selectedWarehouse)
      .map(row => {
        const isScanned = scannedSPs.includes(String(row["SP"]).trim());
        return { ...row, isScanned };
      });
      
    return result.sort((a, b) => (a.isScanned === b.isScanned ? 0 : a.isScanned ? 1 : -1));
  }, [excelData, selectedWarehouse, scannedRecords]);

  // Kiểm tra xem trong các dòng thiếu có dòng nào được phép "Không Xác" không?
  const canSkipAnyMissing = useMemo(() => {
    return missingRowsForLockedLK.some(row => {
        const repairType = String(row["Repair Type"] || '').trim().toUpperCase();
        const typeLK = String(row["Type"] || '').trim().toUpperCase();
        return !(repairType === 'IW' && typeLK !== 'PIN');
    });
  }, [missingRowsForLockedLK]);

  // Theo dõi mã LK chưa hoàn thành (Đang dở dang)
  useEffect(() => {
    // NẾU ĐÃ MỞ KHÓA VƯỢT CẤP TOÀN CỤC -> BỎ QUA HOÀN TOÀN TÍNH NĂNG KHÓA
    if (isGlobalOverrideActive) {
      setLockedMaLK(null);
      setViolationCount(0);
      return;
    }

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
           String(row["Warehouse Name"]).trim() === selectedWarehouse && 
           String(row["Defective material code"]).trim() === currentMaLK &&
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
  }, [scannedRecords, selectedWarehouse, excelData, isGlobalOverrideActive]);

  const stats = useMemo(() => {
    let totals = { IW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, OOW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, KMH: { LCD: 0, MAIN: 0, OTHERS: 0 } };
    let scanned = { IW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, OOW: { LCD: 0, MAIN: 0, OTHERS: 0, total: 0 }, KMH: { LCD: 0, MAIN: 0, OTHERS: 0 } };

    if (!selectedWarehouse) return { totals, scanned };

    const processItem = (row, targetObj, qtyStr, repairStr, typeStr, kmhStr) => {
      const qty = parseInt(qtyStr) || 1;
      const rType = repairStr === 'IW' ? 'IW' : (repairStr === 'OOW' ? 'OOW' : null);
      const tType = (typeStr === 'LCD' || typeStr === 'MAIN') ? typeStr : 'OTHERS';
      const isKMH = String(kmhStr) === "1" || String(kmhStr).includes("KMH");
      
      if (isKMH) {
        targetObj.KMH[tType] += qty;
      }
      
      if (rType && targetObj[rType]) {
        // Tối ưu OOW: Nếu là OOW và là hàng KMH thì tự động trừ ra (không cộng vào tổng OOW)
        if (rType === 'OOW' && isKMH) {
          return;
        }
        targetObj[rType][tType] += qty;
        targetObj[rType].total += qty;
      }
    };

    // TỐI ƯU KMH: Lấy danh sách SP đã scan và được gán Remark KMH để đối chiếu với Total
    const scannedKMH_SPs = new Set();
    currentDisplayedRecords.forEach(record => {
       if ((record.status === "Khớp, Trả Xác LK về" || record.status === "Không xác LK") && String(record.remark).includes("KMH")) {
           scannedKMH_SPs.add(String(record.sp || '').trim().toUpperCase());
       }
    });

    excelData.forEach(row => {
      if (String(row["Warehouse Name"]).trim() === selectedWarehouse) {
        const sp = String(row["SP"] || '').trim().toUpperCase();
        let kmhStr = String(row["KMH"] || '');
        
        // Nếu phiếu này đã được scan và đánh dấu KMH ở mục Vượt Cấp, thì tính tổng nó vào KMH luôn
        if (scannedKMH_SPs.has(sp)) {
            kmhStr = "KMH";
        }

        processItem(row, totals, String(row["Consumed quantity"] || row["Consumed"] || '1'), String(row["Repair Type"] || '').trim().toUpperCase(), String(row["Type"] || '').trim().toUpperCase(), kmhStr);
      }
    });

    currentDisplayedRecords.forEach(record => {
      if (record.status === "Khớp, Trả Xác LK về" || record.status === "Không xác LK") {
        processItem(record, scanned, String(record.slg || '1'), String(record.bhDv || '').trim().toUpperCase(), String(record.phanLoai || '').trim().toUpperCase(), String(record.remark || ''));
      }
    });

    return { totals, scanned };
  }, [excelData, selectedWarehouse, currentDisplayedRecords]);

  const showToast = useCallback((message, type = 'info') => {
    setToast({ message: String(message), type });
    setTimeout(() => setToast(null), 4000);
  }, []);

  const fetchGoogleSheetData = useCallback(async () => {
    setIsLoadingData(true);
    // TỐI ƯU: Nhường luồng hiển thị 100ms để trình duyệt kịp vẽ vòng quay loading trước khi fetch
    await new Promise(resolve => setTimeout(resolve, 100)); 

    try {
      const sheetId = "10zbQkv7f_7nwVawF5lFDdqfpbC23UT1YBjgJ3aZyNWU";
      const url = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
      const response = await fetch(url);
      
      if (!response.ok) throw new Error("Network response was not ok");
      const data = await response.arrayBuffer();

      // TỐI ƯU: Dừng thêm 50ms trước khi chạy thư viện XLSX nặng để chống đứng giao diện
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
        Object.keys(row).forEach(key => cleanRow[String(key).trim()] = row[key]);
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

  // ĐÃ SỬA LỖI AUTO-FOCUS: Bắt focus linh hoạt giữa input chính và input vượt cấp OOW
  useEffect(() => {
    const isModalOpen = pendingQRCheck || showDetailedProgressModal || confirmDialog || isWhDropdownOpen || showForceRemarkModal || warningAlert || showPasswordModal || quickNoXacRow || isUploading || showOOWOverrideInput;
    if (!isModalOpen && !hasErrorLock && !hasConflictLock) {
      if (isOOWOverrideActive && oowScanInputRef.current) {
         oowScanInputRef.current.focus();
         oowScanInputRef.current.select();
      } else if (!isOOWOverrideActive && mainInputRef.current) {
         mainInputRef.current.focus();
         mainInputRef.current.select();
      }
    }
  }, [pendingQRCheck, hasErrorLock, hasConflictLock, scannedRecords.length, showDetailedProgressModal, confirmDialog, selectedWarehouse, isWhDropdownOpen, showForceRemarkModal, warningAlert, showPasswordModal, quickNoXacRow, isUploading, showOOWOverrideInput, isOOWOverrideActive]);

  const handleMainScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    
    const rawScan = String(scanInput).trim();
    if (!rawScan) return;

    if (hasErrorLock) {
      showToast("Hệ thống đang khóa! Vui lòng xóa dòng lỗi để tiếp tục.", "error");
      return;
    }
    if (hasConflictLock) {
      showToast("Hệ thống đang khóa do trùng lặp trạng thái! Hãy xóa dòng Không xác LK bị trùng ở đầu bảng lịch sử.", "error");
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
    matchedRow = excelData.find(row => String(row["SP"] || '') === extractedSPString && String(row["Warehouse Name"] || '') === selectedWarehouse);

    // Nếu không tìm thấy, thử tìm theo QR Code nguyên gốc
    if (!matchedRow) {
       matchedRow = excelData.find(row => String(row["QRCode"] || '').trim() === rawScan && String(row["Warehouse Name"] || '') === selectedWarehouse && rawScan !== '');
       if (matchedRow) {
           extractedSPString = String(matchedRow["SP"] || '').trim();
           isQRFirstScan = true;
       }
    }

    if (!matchedRow) {
       setScannedRecords(prev => [createRecord(rawScan, "Không đúng Xác LK hoặc QR ko tồn tại"), ...prev]);
       setScanInput('');
       return;
    }

    if (lockedMaLK && String(matchedRow["Defective material code"] || '') !== lockedMaLK) {
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

    const isDuplicate = currentDisplayedRecords.some(r => r.status === "Khớp, Trả Xác LK về" && (r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === extractedSPString);
    if (isDuplicate) {
       setScannedRecords(prev => [createRecord(rawScan, "Bị trùng, SP hoặc QR này đã được scan"), ...prev]);
       setScanInput('');
       return;
    }

    const repairType = String(matchedRow["Repair Type"] || '').trim().toUpperCase();
    const type = String(matchedRow["Type"] || '').trim().toUpperCase();

    if (repairType === 'IW' && (type === 'MAIN' || type === 'LCD')) {
      if (!isQRFirstScan) {
         setPendingQRCheck({ type: 'WAITING_FOR_QR', rowData: matchedRow, spScan: rawScan });
      } else {
         setPendingQRCheck({ type: 'WAITING_FOR_SP', rowData: matchedRow, qrScan: rawScan });
      }
      setTimeout(() => { if (secondInputRef.current) secondInputRef.current.focus(); }, 100);
    } else {
      setScannedRecords(prev => [createRecord(rawScan, "Khớp, Trả Xác LK về", matchedRow), ...prev]);
      setTimeout(() => {
        if (mainInputRef.current) {
          mainInputRef.current.focus();
          mainInputRef.current.select();
        }
      }, 50);
    }
    
    setScanInput('');
  }, [scanInput, hasErrorLock, hasConflictLock, excelData, selectedWarehouse, currentDisplayedRecords, lockedMaLK, violationCount, missingRowsForLockedLK, showToast]);

  const handleSecondScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    const secondScan = String(secondScanInput).trim();
    if (!secondScan) return;

    // FIX LỖI TRẮNG MÀN HÌNH: Chặn thao tác dư thừa khi dữ liệu xác thực bước 2 chưa sẵn sàng
    if (!pendingQRCheck) return; 

    if (hasErrorLock) {
       showToast("Hệ thống đang khóa! Xóa lỗi bên dưới để tiếp tục.", "error");
       return;
    }

    const { type, rowData, spScan, qrScan } = pendingQRCheck;

    if (type === 'WAITING_FOR_QR') {
      const expectedQR = String(rowData["QRCode"] || '').trim();
      const isMatched = secondScan === expectedQR;
      const status = isMatched ? "Khớp, Trả Xác LK về" : "QR Tem Xác và LK ko khớp nhau";
      setScannedRecords(prev => [createRecord(spScan, status, rowData, { secondQRScanned: secondScan, expectedQR }), ...prev]);
    } 
    else if (type === 'WAITING_FOR_SP') {
      const spString = String(secondScan).split('_').join('').split(' ').join('').toUpperCase();
      const expectedSP = String(rowData["SP"] || '').trim();
      const isMatched = spString === expectedSP;
      const status = isMatched ? "Khớp, Trả Xác LK về" : "Số phiếu/Mã LK không khớp với QR Tem Xác";

      if (isMatched) {
        const isDuplicate = currentDisplayedRecords.some(r => r.status === "Khớp, Trả Xác LK về" && (r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === spString);
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
    
    // TỐI ƯU AUTO-FOCUS: Bắt lại tiêu điểm chuẩn xác vào ô nhập liệu chính sau khi Modal đóng
    setTimeout(() => {
       if (isOOWOverrideActive && oowScanInputRef.current) {
          oowScanInputRef.current.focus();
          oowScanInputRef.current.select(); // Tự động bôi đen để quét đè mã mới ngay
       } else if (mainInputRef.current) {
          mainInputRef.current.focus();
          mainInputRef.current.select(); // Tự động bôi đen để quét đè mã mới ngay
       }
    }, 150);

  }, [secondScanInput, pendingQRCheck, currentDisplayedRecords, hasErrorLock, isOOWOverrideActive, showToast]);

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
          if (!remarkText || !String(remarkText).trim()) {
             hasError = true;
          } else {
             const selectedRow = missingRowsForLockedLK[idx];
             const rawSP = String(selectedRow["SP"] || '');
             const newRecord = createRecord(rawSP, "Không xác LK", selectedRow);
             newRecord.remark = String(remarkText).trim();
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
    if (!quickNoXacReason || !String(quickNoXacReason).trim()) {
      showToast("Vui lòng điền lý do báo mất linh kiện!", "warning");
      return;
    }
    const rawSP = String(quickNoXacRow["SP"] || '');
    const newRecord = createRecord(rawSP, "Không xác LK", quickNoXacRow);
    newRecord.remark = String(quickNoXacReason).trim();

    setScannedRecords(prev => [newRecord, ...prev]);
    setQuickNoXacRow(null);
    setQuickNoXacReason('');
    showToast(`Đã ghi nhận Không xác LK cho số RO: ${newRecord.soRO}`, "success");
  };

  // --- CÁC HÀM XỬ LÝ VƯỢT CẤP OOW (KMH) ---
  const handleOOWOverrideSubmit = useCallback(() => {
    if (oowOverridePassword === 'kk134') {
       setIsOOWOverrideActive(true);
       setShowOOWOverrideInput(false);
       setOowOverridePassword('');
       showToast("Đã mở khóa quét KMH cho nhóm OOW - LCD!", "success");
       setTimeout(() => {
         if (oowScanInputRef.current) {
           oowScanInputRef.current.focus();
           oowScanInputRef.current.select();
         }
       }, 100);
    } else {
       setShowOOWOverrideInput(false);
       setOowOverridePassword('');
    }
  }, [oowOverridePassword, showToast]);

  const handleOOWScan = useCallback((e) => {
    if (e.key !== 'Enter') return;
    const rawScan = String(oowScanInput).trim();
    if (!rawScan) return;

    if (hasErrorLock) {
      showToast("Hệ thống đang khóa! Vui lòng xóa dòng lỗi ở bảng chính để tiếp tục.", "error");
      return;
    }
    if (hasConflictLock) {
      showToast("Hệ thống đang khóa do trùng lặp trạng thái!", "error");
      return;
    }

    let extractedSPString = String(rawScan).split('_').join('').split(' ').join('').toUpperCase();
    let matchedRow = excelData.find(row => String(row["SP"] || '') === extractedSPString && String(row["Warehouse Name"] || '') === selectedWarehouse);

    if (!matchedRow) {
       matchedRow = excelData.find(row => String(row["QRCode"] || '').trim() === rawScan && String(row["Warehouse Name"] || '') === selectedWarehouse && rawScan !== '');
       if (matchedRow) {
           extractedSPString = String(matchedRow["SP"] || '').trim();
       }
    }

    if (!matchedRow) {
       showToast("Không đúng Xác LK hoặc QR ko tồn tại!", "error");
       setOowScanInput('');
       return;
    }

    const repairType = String(matchedRow["Repair Type"] || '').trim().toUpperCase();
    const type = String(matchedRow["Type"] || '').trim().toUpperCase();

    if (repairType !== 'OOW' || type !== 'LCD') {
       showToast("Lỗi: Tính năng này chỉ áp dụng cho mã LK OOW và loại LCD!", "error");
       setOowScanInput('');
       return;
    }

    const isDuplicate = currentDisplayedRecords.some(r => r.status === "Khớp, Trả Xác LK về" && (r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase()) === extractedSPString);
    if (isDuplicate) {
       showToast("Lỗi: Mã này đã được scan rồi!", "error");
       setOowScanInput('');
       return;
    }

    const newRecord = createRecord(rawScan, "Khớp, Trả Xác LK về", matchedRow);
    newRecord.remark = "KMH"; // Đánh dấu mặc định là KMH theo yêu cầu
    
    setScannedRecords(prev => [newRecord, ...prev]);
    setOowScanInput('');
    showToast(`Đã quét KMH thành công: ${newRecord.soRO}`, "success");
    setTimeout(() => {
      if (oowScanInputRef.current) {
        oowScanInputRef.current.focus();
        oowScanInputRef.current.select();
      }
    }, 50);

  }, [oowScanInput, hasErrorLock, hasConflictLock, excelData, selectedWarehouse, currentDisplayedRecords, showToast]);

  // NÚT BẤM: XÁC NHẬN MỞ KHÓA VƯỢT CẤP TẠI BẢNG CẢNH BÁO
  const handleOverrideSubmit = () => {
    if (overridePassword === 'kk134') {
      setIsGlobalOverrideActive(true); // Bật cờ vượt cấp toàn hệ thống
      setLockedMaLK(null);
      setViolationCount(0);
      setShowForceRemarkModal(false);
      setOverridePassword('');
      setShowOverrideInput(false);
      showToast("Đã mở khóa vượt cấp thành công! Bạn có thể scan tự do các mã LK.", "success");
    } else {
      // Nhập sai pass -> Không phản ứng (Silent fail)
      setOverridePassword('');
      setShowOverrideInput(false);
    }
  };

  // ----------------------------------------------------------------------
  // VALIDATION: Kiểm tra trước khi Lưu/Xuất/Đẩy Drive
  // ----------------------------------------------------------------------
  const canExecuteCenterAction = useCallback(() => {
    if (!selectedWarehouse) {
      showToast("Vui lòng chọn Trung Tâm Bảo Hành!", "warning");
      return false;
    }
    if (hasConflictLock) {
      showToast("Hệ thống đang có phiếu trùng lặp trạng thái (Khớp & Không xác LK). Vui lòng XÓA dòng 'Không xác LK' bị trùng ở đầu lịch sử trước!", "error");
      return false;
    }
    if (currentDisplayedRecords.length === 0) {
      showToast("Không có lịch sử quét nào! Bạn chưa thể lưu phiên hay đẩy báo cáo.", "warning");
      return false;
    }
    // Ràng buộc: Phải có ít nhất 1 mã quét thực tế hợp lệ, 
    // không cho phép đẩy lên nếu tất cả đều là "Không xác LK"
    const hasValidScan = currentDisplayedRecords.some(r => r.status === "Khớp, Trả Xác LK về");
    if (!hasValidScan) {
      showToast("Lịch sử chỉ có phiếu 'Không xác LK'. Vui lòng quét ít nhất 1 linh kiện thực tế để có thể lưu phiên/đẩy báo cáo!", "error");
      return false;
    }
    return true;
  }, [selectedWarehouse, currentDisplayedRecords, hasConflictLock, showToast]);

  // HÀM ĐỒNG BỘ TRỰC TIẾP LÊN GOOGLE DRIVE (Thêm biến showOverlay để chạy ngầm)
  const executeUploadToDrive = async (workbook, showOverlay = true) => {
    if (!gasUrl || !String(gasUrl).trim()) return false;
    
    // Chỉ hiện Loading Overlay khi được cho phép
    if (showOverlay) setIsUploading(true);
    
    try {
      // TỐI ƯU: Dừng 100ms để React vẽ Loading UI trước khi bị hàm XLSX.write block luồng chính
      if (showOverlay) await new Promise(resolve => setTimeout(resolve, 100));

      const base64Data = window.XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
      const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
      const filename = `BaoCao_XacLK_${folderName}.xlsx`;

      // Gửi dưới dạng chuỗi text thuần túy để tránh lỗi CORS chặn từ trình duyệt
      const response = await fetch(String(gasUrl), {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify({
          warehouseCode: String(folderName),
          filename: String(filename),
          fileData: String(base64Data)
        })
      });

      const textRes = await response.text();
      try {
        const res = JSON.parse(textRes);
        return !!res.success;
      } catch (parseError) {
        console.error("Lỗi phân tích JSON từ Apps Script:", textRes);
        return false;
      }
    } catch (err) {
      console.error("Lỗi Fetch API:", err);
      return false;
    } finally {
      if (showOverlay) setIsUploading(false);
    }
  };

  // NÚT BẤM: LƯU PHIÊN (Và tự động đẩy lên Drive CHẠY NGẦM)
  const handleSaveSession = async () => {
    if (!canExecuteCenterAction()) return;

    // 1. Tải file JSON xuống máy cục bộ (Thực hiện ngay lập tức không bị khựng)
    const dataStr = JSON.stringify(scannedRecords, null, 2);
    const blob = new Blob([dataStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    
    link.href = url;
    link.download = `PhienScan_${selectedWarehouse ? String(selectedWarehouse).split(' ').join('') : 'All'}_${new Date().toISOString().slice(0, 10).split('-').join('')}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast("Đã tải xuống file sao lưu phiên làm việc. (Đang đồng bộ ngầm báo cáo lên Drive...)", "success");

    // 2. Chạy tiến trình đẩy báo cáo Excel lên Drive chạy nền (Truyền false để KHÔNG hiện Overlay Loading)
    if (gasUrl && String(gasUrl).trim()) {
      if (!window.XLSX) return;
      const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
      const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
      
      const uploadSuccess = await executeUploadToDrive(workbook, false); // Tham số false = Chạy ngầm
      if (uploadSuccess) {
        showToast(`Đã đồng bộ ngầm báo cáo lên Google Drive [${folderName}] thành công!`, "success");
      } else {
        showToast("Đồng bộ ngầm lên Drive gặp lỗi mạng!", "warning");
      }
    }
  };

  // NÚT BẤM: ĐẨY LÊN DRIVE (Cố tình đẩy thủ công -> HIỆN LOADING)
  const handleUploadToDrive = async () => {
    if (!canExecuteCenterAction()) return;
    if (!window.XLSX) {
      showToast("Thư viện Excel đang được tải, vui lòng chờ trong giây lát.", "warning");
      return;
    }

    const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
    const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
    const success = await executeUploadToDrive(workbook, true); // Tham số true = Hiện Loading Overlay
    
    if (success) {
      showToast(`Đẩy lên Google Drive thành công vào thư mục [${folderName}]!`, "success");
    } else {
      showToast("Đồng bộ thất bại! Hãy kiểm tra Console lỗi mạng hoặc quyền truy cập Apps Script.", "error");
    }
  };

  // NÚT BẤM: TẢI EXCEL TTBH VỀ MÁY (Kèm đồng bộ ngầm)
  const handleExportExcel = async () => {
    if (!canExecuteCenterAction()) return;
    if (!window.XLSX) {
      showToast("Thư viện Excel đang được tải, vui lòng thử lại.", "warning");
      return;
    }

    const workbook = generateWorkbookData(selectedWarehouse, currentDisplayedRecords, detailedProgressList);
    const dateStr = new Date().toISOString().slice(0, 10).split('-').join('');
    
    window.XLSX.writeFile(workbook, `BaoCao_ScanLK_${selectedWarehouse ? String(selectedWarehouse).split(' ').join('') : 'All'}_${dateStr}.xlsx`);
    showToast("Đã tải xuống file báo cáo Excel.", "success");

    // Đẩy lên drive ngầm (Truyền false để ẩn Overlay)
    if (gasUrl && String(gasUrl).trim()) {
      const folderName = currentScCode || selectedWarehouse || "TTBH_Unknown";
      const uploadSuccess = await executeUploadToDrive(workbook, false);
      if (uploadSuccess) {
        showToast(`Đã đồng bộ ngầm báo cáo lên Drive [${folderName}]!`, "success");
      }
    }
  };

  // NẠP PHIÊN LÀM VIỆC CŨ: THUẬT TOÁN GỘP THÔNG MINH, PHÁT HIỆN TRÙNG KHÁC TRẠNG THÁI ĐỂ XỬ LÝ LỖI
  const handleLoadSession = useCallback((e) => {
    const inputTarget = e.target;
    const file = inputTarget.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const loadedData = JSON.parse(evt.target.result);
        if (!Array.isArray(loadedData)) throw new Error("Invalid format");

        // Đồng bộ danh sách các nhà kho có trong file
        const loadedWarehouses = [...new Set(loadedData.map(r => String(r.ttbh || "").trim()).filter(Boolean))];
        setWarehouses(prev => {
          const combined = [...new Set([...prev, ...loadedWarehouses])];
          return combined.sort((a, b) => String(a).localeCompare(String(b), 'vi'));
        });

        if (loadedWarehouses.length === 1 && selectedWarehouse !== loadedWarehouses[0]) {
          setSelectedWarehouse(String(loadedWarehouses[0]));
        } else if (!selectedWarehouse && loadedWarehouses.length > 0) {
          setSelectedWarehouse(String(loadedWarehouses[0]));
        }

        // TIẾN HÀNH THUẬT TOÁN GỘP THÔNG MINH PHÁT HIỆN XUNG ĐỘT TRẠNG THÁI
        setScannedRecords(prevOnScreenRecords => {
          const onScreenMap = new Map();
          prevOnScreenRecords.forEach(r => {
            const uniqueKey = `${String(r.rawScan).toUpperCase()}|${String(r.ttbh).trim()}`;
            onScreenMap.set(uniqueKey, r);
          });

          let addCount = 0;
          let skipCount = 0;
          let conflictCount = 0;
          const mergedList = [...prevOnScreenRecords]; // Giữ lại toàn bộ dữ liệu phiên quét hiện tại trên màn hình

          loadedData.forEach(loadedItem => {
            const cleanItem = {
              ...loadedItem,
              sp: loadedItem.sp ? String(loadedItem.sp) : String(loadedItem.rawScan || '').split('_').join('').split(' ').join('').toUpperCase()
            };

            const uniqueKey = `${String(cleanItem.rawScan).toUpperCase()}|${String(cleanItem.ttbh).trim()}`;

            if (onScreenMap.has(uniqueKey)) {
              const onScreenItem = onScreenMap.get(uniqueKey);
              if (String(onScreenItem.status) === String(cleanItem.status)) {
                // Trùng khớp hoàn toàn cả về mã quét và trạng thái đối chiếu -> Bỏ qua không nạp trùng lặp dòng
                skipCount++;
              } else {
                // BỊ TRÙNG MÃ QUÉT NHƯNG KHÁC TRẠNG THÁI (Một bên Khớp xác thực, một bên Không xác LK)
                // -> Nạp đè vào hệ thống để cơ chế Conflict Tracker phát hiện xung đột và khóa thao tác buộc trung tâm xử lý
                mergedList.push(cleanItem);
                addCount++;
                conflictCount++;
              }
            } else {
              // Phiếu chưa tồn tại trên màn hình -> Tiến hành nạp bình thường
              mergedList.push(cleanItem);
              addCount++;
            }
          });

          if (conflictCount > 0) {
            showToast(`Phát hiện ${conflictCount} phiếu bị xung đột trạng thái (Vừa có Khớp vừa có Không xác)! Hệ thống đã đẩy lên đầu lịch sử và KHÓA CHỜ xử lý xóa.`, "error");
          } else if (addCount > 0) {
            showToast(`Đã gộp thành công ${addCount} dòng từ file nạp vào phiên quét hiện tại! (Bỏ qua ${skipCount} dòng trùng lặp)`, "success");
          } else {
            showToast(`Toàn bộ ${skipCount} dòng trong file nạp đều đã được bắn sẵn trên màn hình.`, "warning");
          }

          return mergedList;
        });

      } catch (err) {
        showToast("Lỗi định dạng tệp file phiên, không thể đọc dữ liệu!", "error");
      } finally {
        inputTarget.value = null; 
      }
    };
    reader.readAsText(file);
  }, [selectedWarehouse, showToast]);

  // HÀM XUẤT THỰC TẾ CHO NÚT "TẢI TẤT CẢ" (Chạy sau khi nhập đúng Password)
  const executeExportAllExcel = useCallback(() => {
    const workbook = window.XLSX.utils.book_new();
    const globalScannedSPs = scannedRecords
      .filter(r => r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK")
      .map(r => r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase());

    const globalScannedKMH_SPs = new Set();
    scannedRecords.forEach(r => {
        if ((r.status === "Khớp, Trả Xác LK về" || r.status === "Không xác LK") && String(r.remark).includes("KMH")) {
            globalScannedKMH_SPs.add(r.sp || String(r.rawScan).split('_').join('').split(' ').join('').toUpperCase());
        }
    });

    const globalDetailedList = excelData.map(row => {
      const isScanned = globalScannedSPs.includes(String(row["SP"] || "").trim());
      return { ...row, isScanned };
    });

    const statsData = [];
    warehouses.forEach(wh => {
      const whData = globalDetailedList.filter(r => String(r["Warehouse Name"] || "").trim() === String(wh));
      if (whData.length === 0) return;

      const iwData = whData.filter(r => String(r["Repair Type"]).toUpperCase() === 'IW');
      const iwGroups = {};
      iwData.forEach(r => {
        const t = String(r["Type"]).toUpperCase();
        const type = (t === 'LCD' || t === 'MAIN') ? t : 'OTHERS';
        const key = `${type}|${String(r["Defective material code"])}|${String(r["Product Name"])}`;
        if (!iwGroups[key]) iwGroups[key] = { type, maLK: String(r["Defective material code"] || ''), tenLK: String(r["Product Name"] || ''), total: 0, scanned: 0 };
        iwGroups[key].total += 1;
        if (r.isScanned) iwGroups[key].scanned += 1;
      });

      Object.values(iwGroups).forEach(g => {
        statsData.push({ "TTBH": String(wh), "Phân Loại": "IW", "Nhóm LK": g.type, "Mã LK": g.maLK, "Tên LK": g.tenLK, "Cần Scan": g.total, "Đã Scan": g.scanned, "Còn Thiếu": g.total - g.scanned });
      });

      const isRowKMH = (r) => {
          const sp = String(r["SP"] || '').trim().toUpperCase();
          return String(r["KMH"]) === "1" || String(r["KMH"]).includes("KMH") || globalScannedKMH_SPs.has(sp);
      };

      const oowData = whData.filter(r => String(r["Repair Type"]).toUpperCase() === 'OOW' && !isRowKMH(r));
      const oowGroups = { 'LCD': { total: 0, scanned: 0 }, 'MAIN': { total: 0, scanned: 0 }, 'OTHERS': { total: 0, scanned: 0 } };
      oowData.forEach(r => {
        const t = String(r["Type"]).toUpperCase();
        const type = (t === 'LCD' || t === 'MAIN') ? t : 'OTHERS';
        oowGroups[type].total += 1;
        if (r.isScanned) oowGroups[type].scanned += 1;
      });
      ['LCD', 'MAIN', 'OTHERS'].forEach(type => {
        if (oowGroups[type].total > 0) statsData.push({ "TTBH": String(wh), "Phân Loại": "OOW", "Nhóm LK": type, "Mã LK": "-", "Tên LK": `Tổng số lượng ${type} (OOW)`, "Cần Scan": oowGroups[type].total, "Đã Scan": oowGroups[type].scanned, "Còn Thiếu": oowGroups[type].total - oowGroups[type].scanned });
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
        if (otherGroups[type].total > 0) statsData.push({ "TTBH": String(wh), "Phân Loại": "Khác", "Nhóm LK": type, "Mã LK": "-", "Tên LK": `Tổng số lượng ${type} (Khác)`, "Cần Scan": otherGroups[type].total, "Đã Scan": otherGroups[type].scanned, "Còn Thiếu": otherGroups[type].total - otherGroups[type].scanned });
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
        const slg = parseInt(r["Consumed quantity"] || r["Consumed"] || 1);
        
        const sp = String(r["SP"] || '').trim().toUpperCase();
        const isKMH = String(r["KMH"]) === "1" || String(r["KMH"]).includes("KMH") || globalScannedKMH_SPs.has(sp);

        if (repairType.toUpperCase() === 'OOW' && isKMH) return; // Loại trừ KMH khỏi OOW

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
        "Model": String(g.model),
        "Type": String(g.type),
        "BH/DV": String(g.repairType),
        "Tổng Cần Scan (Slg)": Number(g.totalSlg),
        "Đã Scan Khớp (Slg)": Number(g.scannedSlg),
        "Còn Thiếu (Slg)": Number(g.totalSlg - g.scannedSlg)
    })).sort((a, b) => String(a.Model).localeCompare(String(b.Model)));

    const wsModelStats = window.XLSX.utils.json_to_sheet(modelStatsData);
    wsModelStats['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 20 }];
    window.XLSX.utils.book_append_sheet(workbook, wsModelStats, "ThongKe_TheoModel");

    // ĐỒNG NHẤT CỘT DỮ LIỆU ĐÃ SCAN (ALL)
    const exportScannedData = scannedRecords.map(record => ({
      "Trạng thái": String(record.status), 
      "Cột SP": String(record.sp || record.rawScan), 
      "SC Code": String(record.scCode), 
      "Warehouse Name": String(record.ttbh), 
      "Số RO": String(record.soRO), 
      "BH/DV": String(record.bhDv), 
      "Mã LK": String(record.maLK), 
      "Product Name": String(record.tenLK), 
      "Model": String(record.model),
      "Type": String(record.phanLoai), 
      "Slg": String(record.slg), 
      "Remark": String(record.remark)
    }));
    const wsScanned = window.XLSX.utils.json_to_sheet(exportScannedData);
    wsScanned['!cols'] = [{ wch: 25 }, { wch: 45 }, { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 40 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 15 }];
    window.XLSX.utils.book_append_sheet(workbook, wsScanned, "LichSu_DaScan_TatCa");

    // ĐỒNG NHẤT CỘT DỮ LIỆU CHƯA SCAN (ALL)
    const wsUnscannedData = globalDetailedList.filter(r => !r.isScanned).map(record => ({
      "Trạng thái": "Chưa Scan", 
      "Cột SP": String(record["SP"] || ""), 
      "SC Code": String(record["SC Code"] || record["SC code"] || ''), 
      "Warehouse Name": String(record["Warehouse Name"] || ""), 
      "Số RO": String(record["After-sales work order No."] || ""), 
      "BH/DV": String(record["Repair Type"] || ""), 
      "Mã LK": String(record["Defective material code"] || ""), 
      "Product Name": String(record["Product Name"] || ""), 
      "Model": String(record["Product Model"] || record["Model"] || ''), 
      "Type": String(record["Type"] || ""), 
      "Slg": String(record["Consumed quantity"] || record["Consumed"] || ''), 
      "Remark": String(getRemark(record))
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
    const s = String(status);
    if (s === "Khớp, Trả Xác LK về") return <span className="inline-flex items-center text-emerald-600 font-bold text-sm"><CheckCircle className="w-5 h-5 mr-1.5" /> Khớp, Trả Xác</span>;
    if (s === "Không xác LK") return <span className="inline-flex items-center text-gray-600 font-bold bg-gray-100 border border-gray-300 px-3 py-1.5 rounded-md text-sm"><AlertTriangle className="w-5 h-5 mr-1.5" /> Không xác LK</span>;
    if (s.includes("trùng")) return <span className="inline-flex items-center text-red-600 font-bold bg-red-100 px-3 py-1.5 rounded-md text-sm"><Ban className="w-5 h-5 mr-1.5" /> Lỗi trùng lặp</span>;
    if (s.includes("định dạng") || s.includes("không tồn tại")) return <span className="inline-flex items-center text-red-600 font-bold bg-red-100 px-3 py-1.5 rounded-md text-sm"><AlertTriangle className="w-5 h-5 mr-1.5" /> Sai định dạng mã</span>;
    return <span className="inline-flex items-center text-rose-500 font-bold bg-rose-50 px-3 py-1.5 rounded-md text-sm"><XCircle className="w-5 h-5 mr-1.5" /> Sai linh kiện / QR</span>;
  }, []);

  const renderDetailedTableGroup = useCallback((title, dataList, titleColorClass) => {
    if (dataList.length === 0) return null;
    const isOOWGroup = String(title).includes("OOW");

    return (
      <div className="mb-8">
        <div className="flex flex-col lg:flex-row lg:items-center gap-3 mb-4">
          <h3 className={`font-bold text-base flex items-center uppercase ${titleColorClass} m-0`}>
            <span className="w-2 h-6 bg-current mr-2 rounded-sm opacity-80"></span>
            Nhóm phiếu {String(title)} 
            <span className="ml-2 bg-gray-100 text-gray-700 px-3 py-0.5 rounded-full text-sm border border-gray-200">{Number(dataList.length)}</span>
          </h3>
          
          {isOOWGroup && (
            <div className="flex items-center ml-2">
              {!isOOWOverrideActive ? (
                !showOOWOverrideInput ? (
                  <button 
                    onClick={() => setShowOOWOverrideInput(true)} 
                    className="text-xs text-gray-400 hover:text-blue-600 font-semibold underline decoration-dotted underline-offset-2 transition-colors cursor-pointer"
                  >
                    Mở khóa vượt cấp
                  </button>
                ) : (
                  <div className="flex items-center gap-2">
                    <input 
                      type="password" 
                      value={oowOverridePassword} 
                      onChange={e => setOowOverridePassword(e.target.value)} 
                      onKeyDown={e => { if (e.key === 'Enter') handleOOWOverrideSubmit(); }} 
                      placeholder="Nhập pass..." 
                      className="w-28 border border-gray-300 rounded-lg px-2 py-1 text-sm focus:outline-none focus:border-blue-500 bg-white" 
                      autoFocus 
                    />
                    <button onClick={handleOOWOverrideSubmit} className="text-xs bg-blue-600 text-white px-3 py-1.5 rounded-lg font-bold shadow-sm">OK</button>
                    <button onClick={() => { setShowOOWOverrideInput(false); setOowOverridePassword(''); }} className="text-xs text-gray-500 font-bold px-1 hover:text-gray-700">Hủy</button>
                  </div>
                )
              ) : (
                <div className="relative flex items-center ml-2">
                  <ScanLine className="absolute left-3 w-4 h-4 text-blue-600" />
                  <input 
                    ref={oowScanInputRef}
                    autoFocus
                    type="text" 
                    value={oowScanInput} 
                    onChange={e => setOowScanInput(e.target.value)} 
                    onKeyDown={handleOOWScan} 
                    placeholder="Scan trực tiếp mã KMH (OOW-LCD)..." 
                    className="border-2 border-blue-400 bg-blue-50 rounded-xl pl-9 pr-4 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 text-blue-900 font-semibold w-72 shadow-inner" 
                  />
                </div>
              )}
            </div>
          )}
        </div>
        <table className="w-full text-left border-collapse whitespace-nowrap bg-white rounded-lg overflow-hidden shadow-sm border border-gray-200">
          <thead className="bg-gray-100 font-bold text-sm uppercase tracking-wider text-gray-600 border-b border-gray-200">
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
                  <td className="p-3 px-4 w-[300px]">
                    <div className="flex items-center gap-2">
                      {row.isScanned ? (
                        <span className="inline-flex items-center text-gray-500 font-bold bg-gray-100 px-3 py-1.5 rounded text-sm border border-gray-200"><CheckCircle className="w-4 h-4 mr-1.5" /> Đã Scan</span>
                      ) : (
                        <>
                          <span className="inline-flex items-center text-red-500 font-bold bg-red-50 px-3 py-1.5 rounded text-sm border border-red-100"><AlertTriangle className="w-4 h-4 mr-1.5" /> Chưa Scan</span>
                          {allowQuickNoXac && (
                            <button
                              onClick={() => {
                                setQuickNoXacRow(row);
                                setQuickNoXacReason('');
                              }}
                              className="text-xs font-bold bg-red-100 text-red-700 hover:bg-red-200 px-2.5 py-1.5 rounded border border-red-200 transition-colors shadow-sm shrink-0"
                            >
                              Không xác LK
                            </button>
                          )}
                        </>
                      )}
                    </div>
                  </td>
                  <td className="p-3 px-4">
                    <div className="flex items-center gap-2">
                      <span className="font-mono text-[15px] text-gray-700">{String(row["After-sales work order No."] || "")}</span>
                      {row["After-sales work order No."] && (
                        <a href={`https://gcsm-sg.oppoit.com/order/order-management/after-sales-order/${String(row["After-sales work order No."])}/detail`} target="_blank" rel="noopener noreferrer" className="inline-flex items-center text-xs font-bold bg-blue-50 text-blue-600 hover:bg-blue-100 hover:text-blue-700 px-2 py-1 rounded border border-blue-200 transition-colors" title="Kiểm tra hệ thống GCSM">
                          <ExternalLink className="w-3.5 h-3.5 mr-1" /> Check GCSM
                        </a>
                      )}
                    </div>
                  </td>
                  <td className="p-3 px-4 font-mono text-[15px] text-gray-700">{String(row["Defective material code"] || "")}</td>
                  <td className="p-3 px-4 text-gray-700 max-w-[280px] truncate font-semibold text-[15px]" title={String(row["Product Name"] || "")}>{String(row["Product Name"] || "")}</td>
                  <td className="p-3 px-4 text-gray-600 text-[15px]">{String(row["Product Model"] || row["Model"] || '')}</td>
                  <td className="p-3 px-4 text-gray-500 text-sm font-medium">{String(row["Type"] || "")}</td>
                  <td className="p-3 px-4 font-bold text-gray-800 text-center text-[15px]">{String(row["Consumed quantity"] || row["Consumed"] || '1')}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }, [quickNoXacRow, quickNoXacReason, showToast, showOOWOverrideInput, oowOverridePassword, isOOWOverrideActive, oowScanInput, handleOOWOverrideSubmit, handleOOWScan]);

  return (
    <div className="h-screen flex flex-col bg-[#f8f9fa] text-base font-sans text-gray-800 overflow-hidden">
      
      {/* MÀN HÌNH CHỜ (LOADING OVERLAY) TOÀN MÀN HÌNH - ĐÃ TỐI ƯU GIAO DIỆN GỌN GÀNG NHẤT */}
      {(isLoadingData || isUploading) && (
        <div className="fixed inset-0 bg-gray-900/50 backdrop-blur-sm flex items-center justify-center z-[9999] transition-all duration-300">
          <div className="bg-white px-8 py-5 rounded-2xl shadow-2xl flex items-center gap-4 border border-gray-100">
            <Loader2 className="w-8 h-8 text-blue-600 animate-spin" />
            <span className="font-bold text-xl text-gray-800 tracking-tight">Đang tải dữ liệu...</span>
          </div>
        </div>
      )}

      {toast && (
        <div className={`fixed bottom-8 right-8 z-[300] px-6 py-4 rounded-xl shadow-2xl text-white font-bold text-base transition-all duration-300 transform translate-y-0 opacity-100 ${toast.type === 'error' ? 'bg-red-500' : toast.type === 'success' ? 'bg-emerald-500' : toast.type === 'warning' ? 'bg-amber-500' : 'bg-blue-500'}`}>
          {String(toast.message)}
        </div>
      )}

      {/* Warning Alert Modal */}
      {warningAlert && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-md w-full border border-red-100 text-center relative z-50">
            <AlertTriangle className="w-16 h-16 text-amber-500 mx-auto mb-5 animate-bounce" />
            <h3 className="font-bold text-xl text-amber-600 mb-3">CẢNH BÁO (Lần {Number(violationCount)})</h3>
            <p className="text-gray-700 whitespace-pre-line mb-8 font-medium text-base">{String(warningAlert)}</p>
            <button onClick={() => setWarningAlert(null)} className="w-full py-3 text-base bg-amber-500 text-white rounded-xl hover:bg-amber-600 font-bold transition-colors shadow-md">Tôi đã hiểu</button>
          </div>
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-gray-900/40 backdrop-blur-sm flex items-center justify-center z-[150] p-4">
          <div className="bg-white rounded-2xl shadow-2xl p-8 max-w-md w-full border border-gray-100 relative z-50">
            <h3 className="font-bold text-xl mb-3 text-gray-800">{String(confirmDialog.title)}</h3>
            <p className="text-gray-600 mb-8 text-base">{String(confirmDialog.message)}</p>
            <div className="flex justify-end gap-4">
              <button onClick={() => setConfirmDialog(null)} className="px-5 py-2.5 bg-gray-100 text-base rounded-xl hover:bg-gray-200 font-bold transition-colors">Hủy</button>
              <button onClick={() => { confirmDialog.onConfirm(); setConfirmDialog(null); }} className="px-5 py-2.5 bg-red-500 text-white text-base rounded-xl hover:bg-red-600 font-bold transition-colors shadow-md">Đồng ý</button>
            </div>
          </div>
        </div>
      )}

      {/* Password Verification Modal */}
      {showPasswordModal && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-sm flex items-center justify-center z-[250] p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-8 max-w-sm w-full border border-gray-200 relative z-50 text-center">
            <Settings className="w-14 h-14 text-blue-600 mx-auto mb-4 animate-spin" style={{ animationDuration: '4s' }} />
            <h3 className="font-bold text-lg text-gray-800 mb-2">Xác thực xuất dữ liệu</h3>
            <p className="text-sm text-gray-500 mb-6">Vui lòng nhập mật khẩu xác nhận quyền xuất báo cáo TỔNG HỢP toàn quốc.</p>
            
            <input 
              ref={passwordInputRef}
              type="password"
              placeholder="Nhập mật khẩu..."
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') handlePasswordSubmit(); }}
              className="w-full text-center border-2 border-gray-200 rounded-xl px-4 py-3 text-base focus:outline-none focus:border-blue-500 bg-gray-50/50 font-mono mb-6"
            />
            
            <div className="flex gap-3">
              <button 
                onClick={() => { setShowPasswordModal(false); setPasswordInput(''); }}
                className="flex-1 py-3 text-sm bg-gray-100 text-gray-600 rounded-xl hover:bg-gray-200 font-bold transition-colors"
              >
                Hủy bỏ
              </button>
              <button 
                onClick={handlePasswordSubmit}
                className="flex-1 py-3 text-sm bg-blue-600 text-white rounded-xl hover:bg-blue-700 font-bold transition-colors shadow-lg"
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
          <div className="bg-white rounded-3xl shadow-2xl p-8 max-w-lg w-full border border-red-100 relative z-[250]">
            <div className="flex items-center text-red-600 mb-5 border-b border-gray-100 pb-4">
              <AlertTriangle className="w-8 h-8 mr-3 animate-pulse" />
              <h2 className="text-xl font-bold text-gray-800">Khai Báo Không Xác Linh Kiện</h2>
            </div>
            
            <div className="bg-red-50 border border-red-200 rounded-2xl p-5 mb-6 space-y-2 text-sm text-red-800">
              <p><strong>Số RO:</strong> {String(quickNoXacRow["After-sales work order No."] || "")}</p>
              <p><strong>Mã LK:</strong> {String(quickNoXacRow["Defective material code"] || "")}</p>
              <p><strong>Tên LK:</strong> {String(quickNoXacRow["Product Name"] || "")}</p>
              <p><strong>Phân Loại:</strong> {String(quickNoXacRow["Repair Type"] || "")} - {String(quickNoXacRow["Type"] || "")}</p>
            </div>

            <div className="mb-8">
              <label className="block text-gray-700 text-sm font-extrabold uppercase mb-3">Nhập lý do báo mất linh kiện (Bắt buộc):</label>
              <input
                type="text"
                value={quickNoXacReason}
                onChange={(e) => setQuickNoXacReason(e.target.value)}
                placeholder="Ví dụ: TTBH báo mất, Khách không trả..."
                className="w-full border-2 border-gray-200 rounded-xl px-5 py-3.5 text-base focus:outline-none focus:border-red-400 bg-gray-50/30 font-semibold"
                autoFocus
              />
            </div>

            <div className="flex justify-end gap-4">
              <button 
                onClick={() => { setQuickNoXacRow(null); setQuickNoXacReason(''); }}
                className="px-6 py-3 text-sm bg-gray-100 text-gray-600 hover:bg-gray-200 rounded-xl font-bold transition-colors"
              >
                Hủy bỏ
              </button>
              <button 
                onClick={handleQuickNoXacSubmit}
                className="px-6 py-3 text-sm bg-red-600 hover:bg-red-700 text-white rounded-xl font-bold transition-colors shadow-lg"
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
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl border border-gray-100 flex flex-col max-h-[90vh] relative z-50">
            <div className="p-6 border-b border-gray-100 shrink-0 bg-rose-50 rounded-t-2xl">
               <h3 className="font-bold text-2xl text-rose-700 flex items-center"><AlertTriangle className="w-8 h-8 mr-3" /> Cảnh báo: Phải Scan xong toàn bộ 1 mã LK của IW hoặc OOW</h3>
               <p className="text-rose-600 mt-3 text-base font-medium">Bạn đã cố tình quét mã linh kiện khác khi <strong>Mã LK: {String(lockedMaLK)}</strong> vẫn chưa được quét xong!</p>
               <p className="text-rose-700 mt-2 text-sm italic font-bold">* Lưu ý: Các loại linh kiện Bảo hành (IW) không phải là PIN bắt buộc phải thu hồi xác, KHÔNG được phép tick chọn "Không xác LK".</p>
            </div>
            
            <div className="p-6 flex-1 overflow-y-auto">
               <p className="font-bold text-gray-800 text-base mb-4">Vui lòng kiểm tra và xác nhận tình trạng các số RO còn thiếu dưới đây:</p>
               <div className="border border-gray-200 rounded-xl overflow-hidden">
                 <table className="w-full text-left text-base whitespace-nowrap">
                   <thead className="bg-gray-50 border-b border-gray-200">
                     <tr>
                       <th className="p-4 px-5 font-bold text-gray-600 w-16 text-center text-sm uppercase">Chọn</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Số RO</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Tên LK</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Loại / Nhóm</th>
                       <th className="p-4 px-5 font-bold text-gray-600 text-sm uppercase">Nhập Lý do (Bắt buộc)</th>
                     </tr>
                   </thead>
                   <tbody className="divide-y divide-gray-100">
                     {missingRowsForLockedLK.map((row, idx) => {
                        const repairType = String(row["Repair Type"] || '').trim().toUpperCase();
                        const typeLK = String(row["Type"] || '').trim().toUpperCase();
                        const isStrictIW = repairType === 'IW' && typeLK !== 'PIN';

                        return (
                          <tr key={idx} className={`transition-colors ${isStrictIW ? 'bg-gray-50/50 opacity-70' : 'hover:bg-gray-50'}`}>
                             <td className="p-3 px-5 text-center">
                               <input
                                  type="checkbox"
                                  disabled={isStrictIW}
                                  checked={!!selectedMissingRows[idx]}
                                  onChange={(e) => setSelectedMissingRows({...selectedMissingRows, [idx]: e.target.checked})}
                                  className="w-5 h-5 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer disabled:cursor-not-allowed"
                               />
                             </td>
                             <td className="p-3 px-5 font-mono font-bold text-[15px] text-gray-700">{String(row["After-sales work order No."] || "")}</td>
                             <td className="p-3 px-5 text-gray-700 truncate max-w-[240px] font-medium" title={String(row["Product Name"] || "")}>{String(row["Product Name"] || "")}</td>
                             <td className="p-3 px-5 text-sm font-bold text-gray-500">
                                {repairType} - {typeLK}
                             </td>
                             <td className="p-3 px-5">
                               {isStrictIW ? (
                                  <span className="text-red-500 text-sm font-bold">Bắt buộc có xác</span>
                               ) : (
                                  <input
                                    type="text"
                                    placeholder="Lý do (TT báo mất...)"
                                    disabled={!selectedMissingRows[idx]}
                                    value={rowRemarks[idx] || ''}
                                    onChange={(e) => setRowRemarks({...rowRemarks, [idx]: e.target.value})}
                                    className="border border-gray-300 rounded-lg px-3 py-2 text-sm w-full focus:border-red-400 focus:outline-none bg-white disabled:bg-gray-100 disabled:cursor-not-allowed"
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

            <div className="p-5 border-t border-gray-100 shrink-0 bg-gray-50 rounded-b-2xl flex justify-between items-center">
              <div className="flex flex-col items-start gap-1">
                <span className="text-sm text-gray-500 italic font-medium">Hệ thống sẽ mở khóa quét mã tiếp theo khi hoàn tất.</span>
                {!showOverrideInput ? (
                  <button
                    onClick={() => setShowOverrideInput(true)}
                    className="text-xs text-gray-400 hover:text-blue-600 font-semibold underline decoration-dotted underline-offset-2 transition-colors cursor-pointer"
                  >
                    Mở khóa vượt cấp
                  </button>
                ) : (
                  <div className="flex items-center gap-2 mt-1">
                    <input
                      type="password"
                      value={overridePassword}
                      onChange={(e) => setOverridePassword(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter') handleOverrideSubmit(); }}
                      placeholder="Nhập mã..."
                      className="w-28 border border-gray-300 rounded-lg px-2 py-1 text-sm focus:outline-none focus:border-blue-500 bg-white"
                      autoFocus
                    />
                    <button
                      onClick={handleOverrideSubmit}
                      className="text-xs bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg font-bold transition-colors shadow-sm"
                    >
                      OK
                    </button>
                    <button
                      onClick={() => { setShowOverrideInput(false); setOverridePassword(''); }}
                      className="text-xs text-gray-500 hover:text-gray-700 font-bold px-1"
                    >
                      Hủy
                    </button>
                  </div>
                )}
              </div>
              <div className="flex gap-4">
                 <button onClick={() => { setShowForceRemarkModal(false); setViolationCount(0); setSelectedMissingRows({}); setRowRemarks({}); setShowOverrideInput(false); setOverridePassword(''); }} className="px-6 py-2.5 bg-gray-200 text-gray-700 rounded-xl hover:bg-gray-300 font-bold text-base transition-colors">Đóng tạm thời</button>
                 {canSkipAnyMissing && (
                    <button onClick={handleBulkForceRemark} className="px-6 py-2.5 bg-red-500 text-white rounded-xl hover:bg-red-600 font-bold text-base transition-colors shadow-md">Xác nhận Không Xác LK</button>
                 )}
              </div>
            </div>
          </div>
        </div>
      )}

      {pendingQRCheck && (
        <div className="fixed inset-0 bg-gray-900/50 backdrop-blur-md flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-3xl shadow-2xl p-8 max-w-lg w-full border border-gray-100">
            <div className="flex items-center text-amber-600 mb-5 border-b border-gray-100 pb-4">
              <AlertTriangle className="w-8 h-8 mr-3" />
              <h2 className="text-xl font-bold text-gray-800">Xác Thực 2 Bước (SP & QR)</h2>
            </div>
            <div className="bg-amber-50 border border-amber-200 rounded-2xl p-5 mb-6">
              <p className="font-bold text-amber-800 text-base">
                {pendingQRCheck.type === 'WAITING_FOR_QR' 
                  ? `Xác ${String(pendingQRCheck.rowData["Repair Type"] || '')} ${String(pendingQRCheck.rowData["Type"] || '')} yêu cầu 100% phải Scan QR Tem Xác LK!`
                  : `Đã nhận mã QR Tem Xác. Vui lòng Scan Số phiếu/Mã LK (SP) để hoàn tất!`}
              </p>
              <p className="text-sm text-gray-700 mt-2 font-mono">Mã Phiếu: <strong className="bg-white px-2 py-0.5 rounded border border-amber-200 text-[15px]">{String(pendingQRCheck.rowData["After-sales work order No."] || "")}</strong></p>
            </div>
            <div className="mb-6">
              <label className="block text-gray-700 text-base font-bold mb-3">
                {pendingQRCheck.type === 'WAITING_FOR_QR' ? 'Quét mã QR trên tem xác:' : 'Quét mã SP (Số Phiếu_Mã LK):'}
              </label>
              <input ref={secondInputRef} autoFocus type="text" maxLength={100} value={secondScanInput} onChange={(e) => setSecondScanInput(e.target.value)} onKeyDown={handleSecondScan} 
                placeholder={pendingQRCheck.type === 'WAITING_FOR_QR' ? 'Bắn mã QR vào đây...' : 'Bắn mã SP vào đây...'} 
                className="w-full border-2 border-amber-300 rounded-xl px-5 py-4 text-lg font-mono focus:outline-none focus:border-amber-500 shadow-inner bg-amber-50/30 font-semibold text-gray-800" />
            </div>
            <div className="flex justify-end">
              <button onClick={() => { setPendingQRCheck(null); setSecondScanInput(''); }} className="px-6 py-3 text-base bg-gray-100 text-gray-600 rounded-xl hover:bg-gray-200 font-bold transition-colors">Hủy bỏ</button>
            </div>
          </div>
        </div>
      )}

      {showDetailedProgressModal && (
        <div className="fixed inset-0 bg-gray-900/60 backdrop-blur-md flex items-center justify-center z-[100] p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col border border-gray-100 relative z-50">
            <div className="flex justify-between items-center p-5 border-b border-gray-100 shrink-0 bg-gray-50 rounded-t-3xl">
              <h2 className="text-xl font-extrabold text-gray-800 flex items-center">
                <ClipboardList className="w-6 h-6 mr-3 text-blue-600" /> Danh sách phiếu CẦN SCAN - <span className="text-blue-700 mx-2">{String(selectedWarehouse)}</span> 
                <span className="text-sm bg-gray-200 text-gray-600 px-3 py-1 rounded-full ml-3 font-bold">Tổng: {Number(detailedProgressList.length)}</span>
              </h2>
              <button onClick={() => setShowDetailedProgressModal(false)} className="p-2 text-gray-400 hover:text-red-500 rounded-xl bg-white border border-gray-200 hover:bg-red-50 transition-colors"><X className="w-6 h-6" /></button>
            </div>
            <div className="overflow-auto flex-1 p-6">
              {detailedProgressList.length === 0 ? (
                <div className="text-center text-gray-500 py-12 font-medium text-lg">TTBH này chưa có dữ liệu yêu cầu đối chiếu.</div>
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

      <header className="bg-white border-b border-gray-200 px-6 py-4 shadow-sm flex justify-between items-center shrink-0 z-20 mt-[2px]">
        <h1 className="text-2xl font-black flex items-center text-gray-800 tracking-tight">
          <FileSpreadsheet className="mr-3 w-7 h-7 text-blue-600" /> Hệ Thống Đối Chiếu Xác Linh Kiện
        </h1>
        <div className="flex items-center text-gray-600 font-semibold">
          {isLoadingData ? (
            <span className="bg-blue-50 text-blue-700 px-4 py-1.5 rounded-full text-sm border border-blue-200 flex items-center font-bold shadow-sm">
              <Loader2 className="w-4 h-4 mr-2 animate-spin" /> Đang đồng bộ ngầm...
            </span>
          ) : (
            <span className="bg-gray-100 px-4 py-1.5 rounded-full text-sm border border-gray-200 flex items-center">
              Tổng dữ liệu gốc: <strong className="text-gray-900 ml-1.5">{Number(excelData.length)}</strong> dòng
            </span>
          )}
        </div>
      </header>

      <main className="flex-1 flex flex-col p-4 gap-4 overflow-hidden w-full max-w-[1700px] mx-auto z-10">
        {/* THÔNG BÁO KHÓA HỆ THỐNG DO TRÙNG LẬP XUNG ĐỘT PHIÊN */}
        {hasConflictLock && (
          <div className="bg-amber-100 border-l-8 border-red-500 p-4 rounded-xl shadow-md flex items-center justify-between animate-pulse">
            <div className="flex items-center gap-4">
              <AlertTriangle className="w-8 h-8 text-red-600" />
              <div>
                <h4 className="font-black text-red-800 text-base">PHÁT HIỆN PHIẾU XUNG ĐỘT TRẠNG THÁI PHIÊN!</h4>
                <p className="text-sm text-red-700 font-semibold mt-1">Một số phiếu vừa quét hợp lệ nhưng tệp phiên nạp lại ghi nhận "Không xác LK". Bắt buộc bấm Xóa (biểu tượng Thùng rác đỏ nhấp nháy) ở đầu bảng lịch sử để mở khóa hệ thống thao tác tiếp tục.</p>
              </div>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 shrink-0">
          <div className="lg:col-span-8 bg-white p-4 rounded-2xl shadow-sm border border-gray-200 flex flex-col relative z-20">
            <div className="flex justify-between items-center mb-3 border-b border-gray-100 pb-3">
              <div className="flex items-center gap-4">
                <h2 className="font-bold text-gray-800 flex items-center text-base">
                  <BarChart2 className="w-5 h-5 mr-2 text-blue-500" /> Thống kê:
                </h2>
                
                {/* Custom Searchable TTBH Combobox */}
                <div className="relative" ref={whDropdownRef}>
                  <button
                    type="button"
                    onClick={() => setIsWhDropdownOpen(!isWhDropdownOpen)}
                    className="border-2 border-blue-200 rounded-xl px-4 py-2 text-sm font-bold bg-blue-50 text-blue-800 focus:outline-none focus:ring-2 focus:ring-blue-400 min-w-[280px] shadow-sm flex items-center justify-between cursor-pointer"
                    disabled={warehouses.length === 0}
                  >
                    <span className="truncate">{String(selectedWarehouse) || "-- Chọn TTBH --"}</span>
                    <Search className="w-4 h-4 ml-2 text-blue-600 opacity-70 shrink-0" />
                  </button>
                  
                  {isWhDropdownOpen && (
                    <div className="absolute left-0 mt-2 w-full bg-white border-2 border-gray-200 rounded-xl shadow-2xl z-50 p-2.5 max-h-[300px] flex flex-col min-w-[300px]">
                      <div className="relative mb-2 flex items-center">
                        <Search className="absolute left-3 w-4 h-4 text-gray-400 pointer-events-none" />
                        <input
                          type="text"
                          placeholder="Gõ tìm nhanh..."
                          value={whSearchQuery}
                          onChange={(e) => setWhSearchQuery(e.target.value)}
                          className="w-full text-sm border border-gray-300 rounded-lg pl-9 pr-3 py-2 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-semibold"
                          autoFocus
                        />
                        {whSearchQuery && (
                          <button 
                            type="button" 
                            onClick={() => setWhSearchQuery('')}
                            className="absolute right-3 text-gray-400 hover:text-gray-700"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                      <div className="overflow-y-auto flex-1 divide-y divide-gray-50">
                        {filteredWarehouses.length === 0 ? (
                          <div className="p-3 text-center text-sm text-gray-400 italic">Không tìm thấy TTBH nào</div>
                        ) : (
                          filteredWarehouses.map((wh, idx) => (
                            <button
                              key={idx}
                              type="button"
                              onClick={() => {
                                setSelectedWarehouse(String(wh));
                                setIsWhDropdownOpen(false);
                                setWhSearchQuery('');
                              }}
                              className={`w-full text-left px-3 py-2.5 text-sm font-bold transition-colors rounded-md flex items-center justify-between ${
                                String(wh) === String(selectedWarehouse) 
                                  ? 'bg-blue-100 text-blue-800' 
                                  : 'hover:bg-gray-100 text-gray-700'
                              }`}
                            >
                              <span className="truncate">{String(wh)}</span>
                              {String(wh) === String(selectedWarehouse) && <CheckCircle className="w-4 h-4 text-blue-600 shrink-0 ml-2" />}
                            </button>
                          ))
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {selectedWarehouse && (
                <button onClick={() => setShowDetailedProgressModal(true)} className="text-sm font-bold bg-gray-50 text-gray-700 px-4 py-2 rounded-xl border border-gray-300 hover:bg-blue-50 hover:text-blue-800 hover:border-blue-300 transition-colors flex items-center shadow-sm">
                  <ClipboardList className="w-4 h-4 mr-2" /> DS Chưa Scan
                </button>
              )}
            </div>

            {!selectedWarehouse ? (
              <div className="text-gray-400 text-center py-8 text-sm italic bg-gray-50 rounded-xl border-2 border-dashed border-gray-200">
                {isLoadingData ? "Đang đồng bộ dữ liệu ngầm, vui lòng chờ trong giây lát..." : "Vui lòng chọn TTBH ở ô phía trên để xem thống kê"}
              </div>
            ) : (
              <div className="flex flex-wrap items-center justify-between gap-6 mt-1">
                <div className="flex-1 min-w-[250px] max-w-lg">
                  <table className="w-full border-collapse text-center text-sm rounded-xl overflow-hidden border border-gray-200">
                    <thead>
                      <tr>
                        <th className="bg-gray-100 p-2 text-gray-700 font-extrabold border-b border-gray-300 w-1/4 uppercase text-xs">Loại</th>
                        <th className="bg-emerald-50 p-2 text-emerald-800 font-extrabold border-b border-emerald-200 w-1/4 uppercase text-xs">IW</th>
                        <th className="bg-blue-50 p-2 text-blue-800 font-extrabold border-b border-blue-200 w-1/4 uppercase text-xs">OOW</th>
                        <th className="bg-rose-50 p-2 text-rose-700 font-extrabold border-b border-rose-200 w-1/4 uppercase text-xs">"KMH"</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 bg-white">
                      {stats.scanned && ['LCD', 'MAIN', 'OTHERS'].map(type => (
                        <tr key={type} className="hover:bg-gray-50">
                          <td className="p-2 font-bold text-gray-700 border-r border-gray-100">{type === 'OTHERS' ? 'Others' : type === 'MAIN' ? 'Main' : type}</td>
                          <td className="p-2 font-bold text-emerald-700 border-r border-gray-100">{Number(stats.scanned.IW[type])} / {Number(stats.totals.IW[type])}</td>
                          <td className="p-2 font-bold text-blue-700 border-r border-gray-100">{Number(stats.scanned.OOW[type])} / {Number(stats.totals.OOW[type])}</td>
                          <td className="p-2 font-bold text-rose-600">{Number(stats.totals.KMH[type]) > 0 ? `${Number(stats.scanned.KMH[type])} / ${Number(stats.totals.KMH[type])}` : '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                
                {/* Horizontal Progress Bars */}
                <div className="flex-1 min-w-[200px] px-5 border-l border-gray-200 flex flex-col justify-center shrink-0">
                  {renderProgressBar(stats.scanned.IW.total, stats.totals.IW.total, 'Tiến độ IW', 'IW')}
                  {renderProgressBar(stats.scanned.OOW.total, stats.totals.OOW.total, 'Tiến độ OOW', 'OOW')}
                </div>
              </div>
            )}
          </div>

          <div className="lg:col-span-4 bg-white p-4 rounded-2xl shadow-sm border border-gray-200 flex flex-col gap-4 justify-center z-10">
            {/* Quản lý Phiên */}
            <div className="flex-1 flex flex-col justify-center bg-gray-50/80 p-3.5 rounded-xl border border-gray-200">
              <h2 className="font-extrabold text-gray-800 mb-3 flex items-center text-xs uppercase tracking-widest">
                <Save className="w-4 h-4 mr-2 text-indigo-500" /> Quản lý Phiên
              </h2>
              <div className="grid grid-cols-2 gap-3">
                <button 
                  onClick={handleSaveSession} 
                  disabled={hasConflictLock}
                  className={`text-sm font-bold py-2.5 rounded-lg flex items-center justify-center shadow-sm transition-colors border-2 ${
                    hasConflictLock 
                      ? 'bg-gray-100 border-gray-300 text-gray-400 cursor-not-allowed' 
                      : 'bg-white border-indigo-300 text-indigo-800 hover:bg-indigo-50 cursor-pointer'
                  }`}
                >
                  <Save className="w-4 h-4 mr-2" /> Lưu Phiên
                </button>
                <label className="text-sm font-bold bg-white border-2 border-amber-300 text-amber-800 py-2.5 rounded-lg hover:bg-amber-50 transition-colors flex items-center justify-center shadow-sm cursor-pointer mb-0">
                  <UploadCloud className="w-4 h-4 mr-2" /> Nạp Phiên
                  <input type="file" accept=".json" onChange={handleLoadSession} className="hidden" />
                </label>
              </div>
            </div>

            {/* Xuất Dữ Liệu */}
            <div className="flex-1 flex flex-col justify-center bg-gray-50/80 p-3.5 rounded-xl border border-gray-200">
              <h2 className="font-extrabold text-gray-800 mb-3 flex items-center text-xs uppercase tracking-widest">
                <Download className="w-4 h-4 mr-2 text-emerald-500" /> Xuất Dữ Liệu
              </h2>
              <div className="grid grid-cols-2 gap-3 mb-3">
                <button 
                  onClick={handleExportExcel}
                  disabled={hasConflictLock}
                  className={`text-sm font-bold py-2.5 rounded-lg flex items-center justify-center shadow-sm transition-colors border-2 ${
                    hasConflictLock 
                      ? 'bg-gray-100 border-gray-300 text-gray-400 cursor-not-allowed' 
                      : 'bg-white border-emerald-300 text-emerald-700 hover:bg-emerald-50 cursor-pointer'
                  }`}
                  title="Tải báo cáo của TTBH đang chọn và tự động đồng bộ Drive"
                >
                  <Download className="w-4 h-4 mr-2" /> Tải TTBH
                </button>
                <button onClick={handleExportAllExcel} className="text-sm font-bold bg-white border-2 border-blue-300 text-blue-700 py-2.5 rounded-lg hover:bg-blue-50 transition-colors flex items-center justify-center shadow-sm" title="Tải báo cáo của tất cả TTBH">
                  <Download className="w-4 h-4 mr-2" /> Tải TẤT CẢ
                </button>
              </div>
              
              {/* NÚT ĐỒNG BỘ DRIVE TRỰC TIẾP THỦ CÔNG */}
              <button 
                onClick={handleUploadToDrive} 
                disabled={isUploading || hasConflictLock}
                className="text-sm font-black bg-emerald-600 text-white border-2 border-emerald-700 py-3 rounded-xl hover:bg-emerald-700 disabled:bg-gray-400 disabled:border-gray-500 disabled:cursor-not-allowed transition-colors flex items-center justify-center shadow-md uppercase tracking-wide"
              >
                {isUploading ? (
                  <>
                    <Loader2 className="w-5 h-5 mr-2 animate-spin" /> Đang đẩy lên Drive...
                  </>
                ) : (
                  <>
                    <UploadCloud className="w-5 h-5 mr-2" /> Đẩy báo cáo lên Drive
                  </>
                )}
              </button>
            </div>
          </div>
        </div>

        <div className={`bg-white rounded-2xl shadow-sm border flex flex-col flex-1 min-h-0 overflow-hidden transition-all duration-300 ${hasErrorLock || hasConflictLock ? 'border-red-400 ring-4 ring-red-100' : 'border-gray-200'} z-0`}>
          <div className={`p-4 border-b flex gap-5 items-center shrink-0 ${hasErrorLock || hasConflictLock ? 'bg-red-50' : 'bg-gray-50'}`}>
            <div className="flex-none font-extrabold text-gray-800 text-base flex items-center">
              Lịch sử đối chiếu
              <span className="bg-white border border-gray-300 text-gray-700 px-3 py-1 rounded-full text-sm ml-3 font-mono shadow-sm">{Number(currentDisplayedRecords.length)}</span>
            </div>
            <div className="flex-1 relative flex items-center max-w-2xl ml-5">
               {hasErrorLock || hasConflictLock ? (
                <span className="text-red-700 font-black text-sm bg-red-200 px-3 py-1 rounded flex items-center shadow-sm shrink-0 border border-red-300">
                  <AlertTriangle className="w-4 h-4 mr-2 animate-bounce" /> HỆ THỐNG ĐANG KHÓA
                </span>
               ) : (
                  <span className="absolute left-4 text-gray-400 z-10"><ScanLine className="w-5 h-5" /></span>
               )}
               <input 
                 ref={mainInputRef}
                 type="text" 
                 maxLength={100}
                 value={scanInput}
                 onChange={(e) => setScanInput(e.target.value)}
                 onKeyDown={handleMainScan}
                 placeholder={isLoadingData ? "Đang tải dữ liệu ngầm..." : hasConflictLock ? "Hãy xóa dòng 'Không xác LK' bị trùng ở đầu lịch sử..." : hasErrorLock ? "Xóa dòng đỏ lỗi bên dưới..." : "Bắn mã vạch vào đây (Tối đa 100 ký tự)..."}
                 disabled={!selectedWarehouse || hasErrorLock || hasConflictLock || isLoadingData}
                 className={`w-full rounded-xl pl-11 pr-5 py-3 text-base font-mono shadow-inner transition-all focus:outline-none border-2 ${
                   hasErrorLock || hasConflictLock ? 'border-red-400 bg-red-100 text-red-800 cursor-not-allowed placeholder-red-600/70 font-bold' : isLoadingData ? 'bg-gray-100 border-gray-200 cursor-not-allowed text-gray-400' : 'border-blue-300 bg-white focus:border-blue-500 focus:ring-4 focus:ring-blue-100 text-gray-900 font-semibold'
                 }`}
               />
            </div>
            <div className="flex-none ml-auto">
              <button 
                onClick={() => setConfirmDialog({
                  title: "Xác nhận xóa Lịch sử",
                  message: "Bạn có chắc chắn muốn xóa toàn bộ lịch sử scan của TTBH này không?",
                  onConfirm: () => { setScannedRecords(prev => prev.filter(r => String(r.ttbh) !== String(selectedWarehouse) && String(r.ttbh) !== '')); showToast("Đã làm mới danh sách thành công.", "success"); }
                })} 
                className="text-sm font-bold bg-white border border-gray-300 text-gray-600 px-4 py-2 rounded-xl hover:bg-red-50 hover:text-red-700 hover:border-red-300 transition-colors shadow-sm flex items-center"
              >
                <Trash2 className="w-4 h-4 mr-2" /> Làm mới danh sách
              </button>
            </div>
          </div>
          
          <div className="flex-1 overflow-y-auto">
            <table className="w-full text-left border-collapse whitespace-nowrap relative">
              <thead className="bg-white sticky top-0 z-10 shadow-sm border-b border-gray-200">
                <tr>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase w-[200px]">Cột Scan QR</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase w-[220px]">Trạng Thái</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">Số RO</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">Mã LK</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">Tên LK</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">Model</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">Loại</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">BH/DV</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase text-center">Slg</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase">Remark</th>
                  <th className="p-3 px-5 border-b border-gray-300 font-extrabold text-gray-700 text-sm uppercase text-center bg-gray-50">Xóa</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100/50">
                {sortedDisplayedRecords.length === 0 ? (
                  <tr>
                    <td colSpan="11" className="p-16 text-center text-gray-400 bg-gray-50/50">
                      <ScanLine className="w-10 h-10 mx-auto mb-4 opacity-30" /> 
                      <span className="font-semibold text-base">Chưa có dữ liệu scan nào ở TTBH này.</span>
                    </td>
                  </tr>
                ) : (
                  sortedDisplayedRecords.map((record) => {
                    const recordSp = String(record.sp || "").trim().toUpperCase();
                    const isConflictNoXac = conflictRecordsSPs.has(recordSp) && String(record.status) === "Không xác LK";
                    const isErrorRow = String(record.status) !== "Khớp, Trả Xác LK về" && String(record.status) !== "Không xác LK";
                    const repairType = String(record.bhDv || '').toUpperCase();
                    
                    let rowColorClass = "hover:bg-gray-50 transition-colors";
                    if (isConflictNoXac) {
                      rowColorClass = "bg-orange-100/90 border-y-2 border-red-500 font-bold hover:bg-orange-200 animate-pulse";
                    } else if (isErrorRow) {
                      rowColorClass = "bg-red-50/80 hover:bg-red-100/80";
                    } else if (String(record.status) === "Không xác LK") {
                      rowColorClass = "bg-gray-50/50";
                    } else if (repairType === 'IW') {
                      rowColorClass = "bg-emerald-50/40 hover:bg-emerald-100/50";
                    } else if (repairType === 'OOW') {
                      rowColorClass = "bg-blue-50/40 hover:bg-blue-100/50";
                    }

                    return (
                      <tr key={record.id} className={rowColorClass}>
                        <td className="p-3 px-5 font-mono text-[15px] font-semibold text-gray-800">
                          <div className="w-[200px] overflow-hidden text-ellipsis whitespace-nowrap" title={String(record.rawScan)}>{String(record.rawScan)}</div>
                        </td>
                        <td className="p-3 px-5">
                          {isConflictNoXac ? (
                            <span className="inline-flex items-center text-red-800 font-black bg-red-200 border border-red-400 px-3 py-1.5 rounded-lg shadow-sm animate-pulse text-sm">
                              <AlertTriangle className="w-4 h-4 mr-1.5" /> Trùng trạng thái (Xóa)
                            </span>
                          ) : (
                            renderStatus(record.status)
                          )}
                        </td>
                        <td className="p-3 px-5 font-mono text-[15px] font-bold">
                          {record.soRO ? (
                            <a href={`https://gcsm-sg.oppoit.com/order/order-management/after-sales-order/${String(record.soRO)}/detail`} target="_blank" rel="noopener noreferrer" className="text-blue-700 hover:text-blue-900 hover:underline transition-colors" title="Kiểm tra trên GCSM">{String(record.soRO)}</a>
                          ) : (
                            <span className="text-gray-500">-</span>
                          )}
                        </td>
                        <td className="p-3 px-5 font-mono text-[15px] font-semibold text-gray-700">{String(record.maLK)}</td>
                        <td className="p-3 px-5 text-gray-800 max-w-[300px] truncate font-bold text-base" title={String(record.tenLK)}>{String(record.tenLK)}</td>
                        <td className="p-3 px-5 text-gray-700 text-[15px] font-medium">{String(record.model)}</td>
                        <td className="p-3 px-3 text-gray-600 text-sm font-bold">{String(record.phanLoai)}</td>
                        <td className="p-3 px-5 font-black text-sm opacity-80">{String(record.bhDv)}</td>
                        <td className="p-3 px-5 font-black text-base text-gray-900 text-center">{String(record.slg)}</td>
                        <td className="p-3 px-5 font-bold text-orange-600 text-sm">{String(record.remark)}</td>
                        <td className={`p-2 text-center border-l border-gray-200/50 ${isErrorRow || isConflictNoXac ? 'bg-red-100/50' : ''}`}>
                          <button 
                            onClick={() => handleDeleteRecord(record.id)} 
                            className={`p-2 rounded-lg transition-all ${
                              isConflictNoXac 
                                ? 'text-white bg-red-600 hover:bg-red-700 shadow-md scale-110 animate-bounce' 
                                : isErrorRow 
                                  ? 'text-white bg-red-500 hover:bg-red-600 shadow-sm' 
                                  : 'text-gray-500 hover:text-red-600 hover:bg-red-100'
                            }`} 
                            title="Xóa dòng này"
                          >
                            <Trash2 className="w-5 h-5 mx-auto" />
                          </button>
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
