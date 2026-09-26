import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  FileSpreadsheet, 
  Layers, 
  Upload, 
  Download, 
  Database, 
  RefreshCw, 
  Trash2, 
  CheckCircle2, 
  AlertTriangle, 
  Search, 
  Filter, 
  ChevronRight, 
  ChevronLeft, 
  ArrowUpDown, 
  Users, 
  GraduationCap, 
  School, 
  FileText, 
  Check, 
  Copy, 
  ExternalLink,
  Info,
  Sliders,
  Plus,
  ArrowRight,
  ShieldCheck,
  History,
  UserX,
  UserCheck,
  BookOpen,
  Calendar,
  Award,
  FileCheck,
  Eye,
  X,
  Server,
  HardDrive,
  Cloud,
  Clock,
  DownloadCloud,
  Play
} from 'lucide-react';
import { 
  GradeFileSource, 
  UnifiedMergedStudentRecord, 
  GradeMergerSummary, 
  StudentHistoryRecord,
  SupabaseConfig, 
  NextcloudConfig,
  NextcloudFile,
  NextcloudGradeSyncConfig,
  NavigationTab 
} from '../types';
import { 
  parseGradeWorkbookFile, 
  mergeGradeFiles, 
  exportUnifiedRecordsToExcel, 
  exportUnifiedRecordsToCsv, 
  exportUnifiedRecordsToJson, 
  generateSupabaseTableSql,
  isPastStudentStatus,
  matchStandardColumn
} from '../services/gradeMergerEngine';
import { StorageService } from '../services/storage';
import { ApiClient } from '../services/apiClient';
import { NextcloudGradeSyncModal } from './NextcloudGradeSyncModal';

interface MultiGradeMergerViewProps {
  supabaseConfig: SupabaseConfig;
  nextcloudConfig: NextcloudConfig;
  files: NextcloudFile[];
  onNavigate: (tab: NavigationTab) => void;
}

export const MultiGradeMergerView: React.FC<MultiGradeMergerViewProps> = ({
  supabaseConfig,
  nextcloudConfig,
  files: nextcloudFiles,
  onNavigate
}) => {
  // Loaded grade workbooks state
  const [gradeFiles, setGradeFiles] = useState<GradeFileSource[]>([]);
  const [mergedRecords, setMergedRecords] = useState<UnifiedMergedStudentRecord[]>(() => 
    StorageService.getMergedGradeRecords()
  );
  const [referenceColumns, setReferenceColumns] = useState<string[]>([]);
  const [summary, setSummary] = useState<GradeMergerSummary | null>(null);
  
  // Student History / Reference State
  const [studentHistories, setStudentHistories] = useState<StudentHistoryRecord[]>(() => 
    StorageService.getStudentHistories()
  );
  const [activeHistoryStudent, setActiveHistoryStudent] = useState<UnifiedMergedStudentRecord | null>(null);
  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState(false);

  // Server storage state
  const [serverSaveStatus, setServerSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle');
  const [lastServerSaveTime, setLastServerSaveTime] = useState<string | null>(null);
  const [serverSyncMessage, setServerSyncMessage] = useState<string>('');

  // Nextcloud Grade Sync State
  const [isNextcloudModalOpen, setIsNextcloudModalOpen] = useState(false);
  const [ncSyncConfig, setNcSyncConfig] = useState<NextcloudGradeSyncConfig>(() => 
    StorageService.getNextcloudGradeSyncConfig()
  );
  const [isSyncingNow, setIsSyncingNow] = useState(false);
  const [timeUntilNextSync, setTimeUntilNextSync] = useState<string | null>(null);

  // New History Entry Form State
  const [historyForm, setHistoryForm] = useState({
    eventType: 'Left School' as StudentHistoryRecord['eventType'],
    date: new Date().toISOString().split('T')[0],
    academicYear: '2023-2024',
    title: '',
    reason: '',
    tcOrCertNumber: '',
    destinationSchool: '',
    description: '',
    addedBy: 'Admin Staff'
  });

  // UI States
  const [isProcessing, setIsProcessing] = useState(false);
  const [processStatusMessage, setProcessStatusMessage] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState('');
  const [studentCohortTab, setStudentCohortTab] = useState<'ALL' | 'ACTIVE' | 'PAST'>('ALL');
  const [selectedGradeFilter, setSelectedGradeFilter] = useState<string>('ALL');
  const [selectedDivisionFilter, setSelectedDivisionFilter] = useState<string>('ALL');
  const [selectedGenderFilter, setSelectedGenderFilter] = useState<string>('ALL');
  const [selectedStatusFilter, setSelectedStatusFilter] = useState<string>('ALL');
  
  // Sorting & Pagination
  const [sortField, setSortField] = useState<string>('');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);

  // Supabase Push Modal & SQL Modal
  const [isSqlModalOpen, setIsSqlModalOpen] = useState(false);
  const [targetTableName, setTargetTableName] = useState('students_consolidated');
  const [isPushingToSupabase, setIsPushingToSupabase] = useState(false);
  const [pushProgress, setPushProgress] = useState(0);
  const [pushStatusMessage, setPushStatusMessage] = useState<string | null>(null);
  const [copiedSql, setCopiedSql] = useState(false);

  // Drag & drop state
  const [isDraggingOver, setIsDraggingOver] = useState(false);

  // File input ref
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Poll Nextcloud Sync Config and Server Storage updates
  useEffect(() => {
    let isMounted = true;

    async function checkSyncConfig() {
      try {
        const res = await ApiClient.getNextcloudGradeSyncConfig();
        if (isMounted && res.success && res.config) {
          setNcSyncConfig(res.config);
          StorageService.saveNextcloudGradeSyncConfig(res.config);
        }
      } catch (e) {}
    }

    checkSyncConfig();
    const interval = setInterval(checkSyncConfig, 12000);
    return () => {
      isMounted = false;
      clearInterval(interval);
    };
  }, []);

  // Compute countdown to next scheduled sync
  useEffect(() => {
    if (!ncSyncConfig.autoSyncEnabled || !ncSyncConfig.nextScheduledSyncTime) {
      setTimeUntilNextSync(null);
      return;
    }

    const timer = setInterval(() => {
      const targetTime = new Date(ncSyncConfig.nextScheduledSyncTime!).getTime();
      const diff = targetTime - Date.now();

      if (diff <= 0) {
        setTimeUntilNextSync('Syncing now...');
        // Refresh server data
        ApiClient.getGradesServerStorage().then(res => {
          if (res.success && res.exists && res.data?.records) {
            setMergedRecords(res.data.records);
            if (res.data.files) setGradeFiles(res.data.files);
            if (res.data.referenceColumns) setReferenceColumns(res.data.referenceColumns);
            if (res.data.summary) setSummary(res.data.summary);
            if (res.data.lastSaved) setLastServerSaveTime(res.data.lastSaved);
          }
        });
      } else {
        const mins = Math.floor(diff / 60000);
        const secs = Math.floor((diff % 60000) / 1000);
        setTimeUntilNextSync(`${mins.toString().padStart(2, '0')}m ${secs.toString().padStart(2, '0')}s`);
      }
    }, 1000);

    return () => clearInterval(timer);
  }, [ncSyncConfig.autoSyncEnabled, ncSyncConfig.nextScheduledSyncTime]);

  // Handler: Quick Trigger Server-Side Nextcloud Sync
  const handleQuickSyncNow = async () => {
    setIsSyncingNow(true);
    setProcessStatusMessage('Pulling updated student workbooks from Nextcloud WebDAV...');

    try {
      const res = await ApiClient.triggerNextcloudGradeSyncNow();
      if (res.success) {
        if (res.records && res.records.length > 0) {
          setMergedRecords(res.records);
          StorageService.saveMergedGradeRecords(res.records);
        }
        if (res.files && res.files.length > 0) {
          setGradeFiles(res.files);
        }
        if (res.referenceColumns && res.referenceColumns.length > 0) {
          setReferenceColumns(res.referenceColumns);
        }
        if (res.summary) {
          setSummary(res.summary);
        }
        setServerSyncMessage(res.message || 'Nextcloud sync completed successfully.');
        setLastServerSaveTime(new Date().toISOString());
        setServerSaveStatus('saved');
      } else {
        setServerSyncMessage(`Nextcloud sync note: ${res.message || res.error}`);
      }

      // Refresh sync config
      const cfgRes = await ApiClient.getNextcloudGradeSyncConfig();
      if (cfgRes.success && cfgRes.config) {
        setNcSyncConfig(cfgRes.config);
      }
    } catch (e: any) {
      console.error('Quick sync error:', e);
      setServerSyncMessage(`Sync error: ${e.message}`);
    } finally {
      setIsSyncingNow(false);
      setProcessStatusMessage('');
    }
  };

  // Handler: Called when Nextcloud Grade Sync Modal finishes a pull
  const handleNextcloudModalSyncComplete = (result: {
    files: any[];
    records: any[];
    referenceColumns: string[];
    summary: any;
    message: string;
  }) => {
    if (result.records && result.records.length > 0) {
      setMergedRecords(result.records);
      StorageService.saveMergedGradeRecords(result.records);
    }
    if (result.files && result.files.length > 0) {
      setGradeFiles(result.files);
    }
    if (result.referenceColumns && result.referenceColumns.length > 0) {
      setReferenceColumns(result.referenceColumns);
      if (!sortField) {
        setSortField(result.referenceColumns[0]);
      }
    }
    if (result.summary) {
      setSummary(result.summary);
    }
    setServerSyncMessage(result.message);
    setLastServerSaveTime(new Date().toISOString());
    setServerSaveStatus('saved');
  };


  // On mount: Load persisted data from server storage
  useEffect(() => {
    let isMounted = true;
    async function loadFromServer() {
      try {
        const res = await ApiClient.getGradesServerStorage();
        if (isMounted && res.success && res.exists && res.data) {
          if (res.data.records && res.data.records.length > 0) {
            setMergedRecords(res.data.records);
            StorageService.saveMergedGradeRecords(res.data.records);
          }
          if (res.data.files && res.data.files.length > 0) {
            setGradeFiles(res.data.files);
          }
          if (res.data.referenceColumns && res.data.referenceColumns.length > 0) {
            setReferenceColumns(res.data.referenceColumns);
            if (!sortField) {
              setSortField(res.data.referenceColumns[0] || 'Student Name');
            }
          }
          if (res.data.summary) {
            setSummary(res.data.summary);
          }
          if (res.data.histories && res.data.histories.length > 0) {
            setStudentHistories(res.data.histories);
            StorageService.saveStudentHistories(res.data.histories);
          }
          if (res.data.lastSaved) {
            setLastServerSaveTime(res.data.lastSaved);
            setServerSaveStatus('saved');
            setServerSyncMessage(`Loaded from persistent server disk storage (${res.data.records?.length || 0} records)`);
          }
        }
      } catch (e: any) {
        console.warn('Could not load grades from server storage:', e);
      }
    }
    loadFromServer();
    return () => { isMounted = false; };
  }, []);

  // Helper: Persist data safely to Server Disk Storage and LocalStorage
  const persistToServerStorage = async (
    filesToSave: GradeFileSource[],
    recordsToSave: UnifiedMergedStudentRecord[],
    colsToSave: string[],
    summaryToSave: GradeMergerSummary | null,
    historiesToSave: StudentHistoryRecord[]
  ) => {
    setServerSaveStatus('saving');
    try {
      // Save in LocalStorage immediately
      StorageService.saveMergedGradeRecords(recordsToSave);
      StorageService.saveStudentHistories(historiesToSave);

      // Save to Server Storage
      const res = await ApiClient.saveGradesServerStorage({
        files: filesToSave,
        records: recordsToSave,
        referenceColumns: colsToSave,
        summary: summaryToSave,
        histories: historiesToSave
      });

      if (res.success) {
        setServerSaveStatus('saved');
        setLastServerSaveTime(res.lastSaved || new Date().toISOString());
        setServerSyncMessage(`Safely saved ${recordsToSave.length} records to server disk storage.`);
      } else {
        setServerSaveStatus('error');
        setServerSyncMessage(res.error || 'Failed to save to server storage');
      }
    } catch (e: any) {
      console.error('Failed to save to server storage:', e);
      setServerSaveStatus('error');
      setServerSyncMessage(e.message || 'Error communicating with server storage');
    }
  };

  // Handler: Execute Merge preserving ONLY Reference Columns
  const runMerge = (filesToMerge: GradeFileSource[], customCols?: string[]) => {
    if (filesToMerge.length === 0) {
      setMergedRecords([]);
      setSummary(null);
      setReferenceColumns([]);
      persistToServerStorage([], [], [], null, studentHistories);
      return;
    }

    setIsProcessing(true);
    setProcessStatusMessage('Merging workbooks, inspecting division sheets, preserving reference columns...');

    setTimeout(async () => {
      try {
        const { records, summary: mergeSummary, referenceColumns: detectedCols } = mergeGradeFiles(filesToMerge, customCols);
        setMergedRecords(records);
        setSummary(mergeSummary);
        setReferenceColumns(detectedCols);
        if (!sortField && detectedCols.length > 0) {
          setSortField(detectedCols[0]);
        }
        
        await persistToServerStorage(filesToMerge, records, detectedCols, mergeSummary, studentHistories);
      } catch (err: any) {
        console.error('Error merging workbooks:', err);
        setProcessStatusMessage('Error during merge: ' + err.message);
      } finally {
        setIsProcessing(false);
        setProcessStatusMessage('');
      }
    }, 50);
  };

  // Convert arrayBuffer to base64 helper
  const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  };

  // Handler: File Upload from User's Device (supports multiple files)
  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = event.target.files;
    if (!selectedFiles || selectedFiles.length === 0) return;

    setIsProcessing(true);
    setProcessStatusMessage(`Reading and analyzing ${selectedFiles.length} workbook(s)...`);

    try {
      const parsedSources: GradeFileSource[] = [];
      
      for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        setProcessStatusMessage(`Parsing file ${i + 1}/${selectedFiles.length}: ${file.name}...`);
        
        const arrayBuf = await file.arrayBuffer();
        const parsed = await parseGradeWorkbookFile({ name: file.name, data: arrayBuf }, gradeFiles.length + i + 1);
        parsedSources.push(parsed);

        // Also save raw file on server storage
        try {
          const base64Data = arrayBufferToBase64(arrayBuf);
          await ApiClient.saveWorkbookFileOnServer(file.name, base64Data);
        } catch (saveErr) {
          console.warn('Could not save raw workbook copy to server:', saveErr);
        }
      }

      const updatedGradeFiles = [...gradeFiles, ...parsedSources];
      setGradeFiles(updatedGradeFiles);
      runMerge(updatedGradeFiles);
    } catch (err: any) {
      console.error('Failed to parse uploaded files:', err);
      setProcessStatusMessage(`Upload error: ${err.message}`);
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  // Drag & drop handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);

    const droppedFiles = e.dataTransfer.files;
    if (!droppedFiles || droppedFiles.length === 0) return;

    setIsProcessing(true);
    setProcessStatusMessage(`Reading and analyzing ${droppedFiles.length} dropped workbook(s)...`);

    try {
      const parsedSources: GradeFileSource[] = [];
      
      for (let i = 0; i < droppedFiles.length; i++) {
        const file = droppedFiles[i];
        if (!file.name.match(/\.(xlsx|xls|csv|ods|xlsm|xlsb)$/i)) continue;

        setProcessStatusMessage(`Parsing file ${i + 1}/${droppedFiles.length}: ${file.name}...`);
        const arrayBuf = await file.arrayBuffer();
        const parsed = await parseGradeWorkbookFile({ name: file.name, data: arrayBuf }, gradeFiles.length + i + 1);
        parsedSources.push(parsed);

        // Also save raw file on server storage
        try {
          const base64Data = arrayBufferToBase64(arrayBuf);
          await ApiClient.saveWorkbookFileOnServer(file.name, base64Data);
        } catch (saveErr) {
          console.warn('Could not save raw workbook copy to server:', saveErr);
        }
      }

      if (parsedSources.length > 0) {
        const updatedGradeFiles = [...gradeFiles, ...parsedSources];
        setGradeFiles(updatedGradeFiles);
        runMerge(updatedGradeFiles);
      }
    } catch (err: any) {
      console.error('Failed to parse dropped files:', err);
      setProcessStatusMessage(`Drop error: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Handler: Toggle Sheet Inclusion in a file
  const toggleSheetInclusion = (fileId: string, sheetName: string) => {
    const updated = gradeFiles.map(f => {
      if (f.id === fileId) {
        return {
          ...f,
          sheets: f.sheets.map(s => s.sheetName === sheetName ? { ...s, included: !s.included } : s)
        };
      }
      return f;
    });
    setGradeFiles(updated);
    runMerge(updated);
  };

  // Handler: Update Grade Name or Division Name
  const updateGradeFileName = (fileId: string, newGradeName: string) => {
    const updated = gradeFiles.map(f => f.id === fileId ? { ...f, detectedGrade: newGradeName } : f);
    setGradeFiles(updated);
    runMerge(updated);
  };

  const updateSheetDivisionName = (fileId: string, sheetName: string, newDivName: string) => {
    const updated = gradeFiles.map(f => {
      if (f.id === fileId) {
        return {
          ...f,
          sheets: f.sheets.map(s => s.sheetName === sheetName ? { ...s, detectedDivision: newDivName.toUpperCase() } : s)
        };
      }
      return f;
    });
    setGradeFiles(updated);
    runMerge(updated);
  };

  // Handler: Remove a Grade File
  const removeGradeFile = (fileId: string) => {
    const updated = gradeFiles.filter(f => f.id !== fileId);
    setGradeFiles(updated);
    runMerge(updated);
  };

  // Handler: Clear All Data
  const clearAllData = async () => {
    if (window.confirm('Are you sure you want to clear all loaded workbooks and consolidated records? This will also clear server storage.')) {
      setGradeFiles([]);
      setMergedRecords([]);
      setSummary(null);
      setReferenceColumns([]);
      StorageService.saveMergedGradeRecords([]);
      try {
        await ApiClient.clearGradesServerStorage();
        setServerSaveStatus('idle');
        setServerSyncMessage('Server storage cleared.');
      } catch {}
    }
  };

  // Handler: Manual Trigger Save to Server
  const handleManualSaveToServer = async () => {
    await persistToServerStorage(gradeFiles, mergedRecords, referenceColumns, summary, studentHistories);
  };

  // Handler: Student History Form Submission
  const handleAddHistoryRecord = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!activeHistoryStudent) return;

    const studentName = activeHistoryStudent.student_name || activeHistoryStudent['Student Name'] || 'Student';
    const admissionNo = activeHistoryStudent.admission_number || activeHistoryStudent['Admission No'] || activeHistoryStudent.id;
    const grade = activeHistoryStudent.grade || activeHistoryStudent['Grade'] || '';
    const division = activeHistoryStudent.division || activeHistoryStudent['Division'] || '';

    const newHistory: StudentHistoryRecord = {
      id: `hist-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      studentId: activeHistoryStudent.id,
      admissionNo,
      studentName,
      grade,
      division,
      eventType: historyForm.eventType,
      date: historyForm.date,
      academicYear: historyForm.academicYear,
      title: historyForm.title || `${historyForm.eventType} - ${studentName}`,
      reason: historyForm.reason,
      tcOrCertNumber: historyForm.tcOrCertNumber,
      destinationSchool: historyForm.destinationSchool,
      description: historyForm.description,
      addedBy: historyForm.addedBy,
      createdAt: new Date().toISOString()
    };

    const updatedHistories = [newHistory, ...studentHistories];
    setStudentHistories(updatedHistories);
    StorageService.saveStudentHistories(updatedHistories);

    // If event marks the student as Left School or Graduated, update the student status in records
    if (['Left School', 'Graduated', 'Transferred', 'Suspended'].includes(historyForm.eventType)) {
      const updatedRecords = mergedRecords.map(r => {
        if (r.id === activeHistoryStudent.id) {
          const statusCol = referenceColumns.find(c => matchStandardColumn(c) === 'academic_status') || 'Status';
          return {
            ...r,
            [statusCol]: historyForm.eventType,
            academic_status: historyForm.eventType
          };
        }
        return r;
      });
      setMergedRecords(updatedRecords);
      await persistToServerStorage(gradeFiles, updatedRecords, referenceColumns, summary, updatedHistories);
    } else {
      // Save history directly on server
      try {
        await ApiClient.saveStudentHistoryOnServer(newHistory);
      } catch {}
    }

    // Reset Form
    setHistoryForm({
      eventType: 'Left School',
      date: new Date().toISOString().split('T')[0],
      academicYear: '2023-2024',
      title: '',
      reason: '',
      tcOrCertNumber: '',
      destinationSchool: '',
      description: '',
      addedBy: 'Admin Staff'
    });
  };

  // Open history modal for a student
  const openStudentHistoryModal = (student: UnifiedMergedStudentRecord) => {
    setActiveHistoryStudent(student);
    setIsHistoryModalOpen(true);
  };

  // Filtered and Sorted Records
  const filteredRecords = useMemo(() => {
    let list = [...mergedRecords];

    // Filter by Cohort tab (ALL vs ACTIVE vs PAST)
    if (studentCohortTab === 'ACTIVE') {
      list = list.filter(r => !isPastStudentStatus(r.academic_status || r['Status']));
    } else if (studentCohortTab === 'PAST') {
      list = list.filter(r => isPastStudentStatus(r.academic_status || r['Status']));
    }

    // Grade Filter
    if (selectedGradeFilter !== 'ALL') {
      list = list.filter(r => (r.grade || r['Grade'] || '') === selectedGradeFilter);
    }

    // Division Filter
    if (selectedDivisionFilter !== 'ALL') {
      list = list.filter(r => (r.division || r['Division'] || '') === selectedDivisionFilter);
    }

    // Gender Filter
    if (selectedGenderFilter !== 'ALL') {
      list = list.filter(r => (r.gender || r['Gender'] || '').toLowerCase() === selectedGenderFilter.toLowerCase());
    }

    // Status Filter
    if (selectedStatusFilter !== 'ALL') {
      list = list.filter(r => (r.academic_status || r['Status'] || '').toLowerCase() === selectedStatusFilter.toLowerCase());
    }

    // Search Query (across all reference columns)
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(r => {
        return referenceColumns.some(col => {
          const val = r[col];
          return val !== undefined && val !== null && String(val).toLowerCase().includes(q);
        }) || (r.student_name && r.student_name.toLowerCase().includes(q))
           || (r.admission_number && r.admission_number.toLowerCase().includes(q));
      });
    }

    // Sorting
    if (sortField) {
      list.sort((a, b) => {
        const valA = a[sortField] !== undefined ? a[sortField] : (a as any)[sortField.toLowerCase().replace(/ /g, '_')] ?? '';
        const valB = b[sortField] !== undefined ? b[sortField] : (b as any)[sortField.toLowerCase().replace(/ /g, '_')] ?? '';

        if (typeof valA === 'number' && typeof valB === 'number') {
          return sortDirection === 'asc' ? valA - valB : valB - valA;
        }
        return sortDirection === 'asc' 
          ? String(valA).localeCompare(String(valB), undefined, { numeric: true })
          : String(valB).localeCompare(String(valA), undefined, { numeric: true });
      });
    }

    return list;
  }, [
    mergedRecords, 
    studentCohortTab, 
    selectedGradeFilter, 
    selectedDivisionFilter, 
    selectedGenderFilter, 
    selectedStatusFilter, 
    searchQuery, 
    sortField, 
    sortDirection, 
    referenceColumns
  ]);

  // Pagination
  const totalPages = Math.ceil(filteredRecords.length / pageSize) || 1;
  const paginatedRecords = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredRecords.slice(start, start + pageSize);
  }, [filteredRecords, currentPage, pageSize]);

  // Available unique filters from current dataset
  const availableGrades = useMemo(() => {
    const grades = new Set<string>();
    mergedRecords.forEach(r => {
      if (r.grade) grades.add(r.grade);
    });
    return Array.from(grades).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }, [mergedRecords]);

  const availableDivisions = useMemo(() => {
    const divs = new Set<string>();
    mergedRecords.forEach(r => {
      if (r.division) divs.add(r.division);
    });
    return Array.from(divs).sort();
  }, [mergedRecords]);

  // Handler: Sort click
  const handleSort = (colName: string) => {
    if (sortField === colName) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(colName);
      setSortDirection('asc');
    }
  };

  // Push to Supabase Direct with intelligent key detection & zero duplicates
  const handlePushToSupabase = async () => {
    if (!supabaseConfig.url || (!supabaseConfig.serviceKey && !supabaseConfig.anonKey)) {
      setPushStatusMessage('Error: Supabase is not connected. Please enter credentials in the Supabase tab.');
      return;
    }

    if (!mergedRecords || mergedRecords.length === 0) {
      setPushStatusMessage('No student records to push.');
      return;
    }

    setIsPushingToSupabase(true);
    setPushProgress(10);
    setPushStatusMessage(`Analyzing ${mergedRecords.length} records for Supabase sync (cell-level deduplication active)...`);

    try {
      const cleanTable = targetTableName.toLowerCase().replace(/[^a-z0-9_]/g, '_');

      // Generate rows to insert preserving reference columns and identifying fields
      const rowsToInsert = mergedRecords.map(r => {
        const dbRow: Record<string, any> = {};
        
        // Populate from reference columns
        referenceColumns.forEach(col => {
          const dbCol = col.toLowerCase().replace(/[^a-z0-9_]/g, '_');
          dbRow[dbCol] = r[col] !== undefined ? r[col] : null;
        });

        // Ensure key identifying fields exist if present on the record
        if (r.admission_number && !dbRow.admission_number && !dbRow.admission_no) {
          dbRow.admission_number = r.admission_number;
        }
        if (r.student_name && !dbRow.student_name && !dbRow.name) {
          dbRow.student_name = r.student_name;
        }
        if (r.grade && !dbRow.grade) {
          dbRow.grade = r.grade;
        }
        if (r.division && !dbRow.division) {
          dbRow.division = r.division;
        }

        return dbRow;
      });

      // Detect the best unique identifier column present in the payload
      const sampleRow = rowsToInsert[0] || {};
      const candidateKeyNames = [
        'admission_no', 'admission_number', 'adm_no', 'admissionno', 
        'student_id', 'studentid', 'student_number', 'std_id',
        'index_no', 'index_number', 'roll_no', 'roll_number', 
        'reg_no', 'registration_no', 'id'
      ];
      
      let detectedConflictKey: string | undefined = undefined;
      for (const key of candidateKeyNames) {
        if (key in sampleRow && sampleRow[key] !== null && sampleRow[key] !== undefined) {
          detectedConflictKey = key;
          break;
        }
      }

      // If no standard unique key column exists, check referenceColumns for matches
      if (!detectedConflictKey) {
        const admissionRefCol = referenceColumns.find(c => matchStandardColumn(c) === 'admission_number');
        if (admissionRefCol) {
          detectedConflictKey = admissionRefCol.toLowerCase().replace(/[^a-z0-9_]/g, '_');
        }
      }

      const res = await ApiClient.upsertSupabaseRecords(
        supabaseConfig, 
        cleanTable, 
        rowsToInsert, 
        detectedConflictKey,
        (progress: any) => {
          setPushProgress(progress.percentage);
          setPushStatusMessage(progress.message);
        }
      );

      if (!res.success) {
        throw new Error(res.error || 'Failed to push to Supabase');
      }

      const skipped = res.skippedCount ?? 0;
      const updated = res.updatedCount ?? 0;
      const inserted = res.insertedCount ?? 0;
      const total = res.totalRecords ?? mergedRecords.length;

      setPushProgress(100);
      setPushStatusMessage(
        `✓ Push Complete! Processed ${total} student records with ZERO duplicates: ${skipped} identical records skipped, ${updated} in-place cell updates applied, and ${inserted} new records inserted into public.${cleanTable}.`
      );
    } catch (err: any) {
      console.error('Supabase push error:', err);
      setPushStatusMessage(`Supabase push failed: ${err.message}. Ensure the table '${targetTableName}' exists with matching columns.`);
    } finally {
      setIsPushingToSupabase(false);
    }
  };

  // Copy SQL
  const handleCopySql = () => {
    const sql = generateSupabaseTableSql(targetTableName, referenceColumns, mergedRecords);
    navigator.clipboard.writeText(sql);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 2500);
  };

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-16">
      {/* Top Banner & Header */}
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2.5">
              <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-lg border border-indigo-100">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
                  Multi-Workbook Student & Class Consolidator
                  <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200">
                    Unlimited Books & Divisions
                  </span>
                </h1>
                <p className="text-sm text-slate-500">
                  Select and merge any number of Excel files and class division sheets (A, B, C, D...) into one consolidated table with exact reference column fidelity and safe server storage.
                </p>
              </div>
            </div>
          </div>

          {/* Action Buttons & Server Status Indicator */}
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Nextcloud AutoSync Live Badge */}
            <div className="flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-medium bg-sky-50 border-sky-200 text-sky-900">
              <Cloud className="w-3.5 h-3.5 text-sky-600 shrink-0" />
              <span>
                {ncSyncConfig.autoSyncEnabled ? (
                  <span className="flex items-center gap-1.5 font-semibold text-sky-950">
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                    Auto-Sync: Every {ncSyncConfig.syncIntervalMinutes}m
                    {timeUntilNextSync && (
                      <span className="text-sky-700 font-normal">({timeUntilNextSync})</span>
                    )}
                  </span>
                ) : (
                  <span className="text-slate-600 font-medium">Nextcloud Sync: Paused</span>
                )}
              </span>
            </div>

            {/* Nextcloud Sync Now Button */}
            <button
              onClick={handleQuickSyncNow}
              disabled={isProcessing || isSyncingNow}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-sky-100 text-sky-800 hover:bg-sky-200 border border-sky-300 transition-colors shadow-sm disabled:opacity-50"
              title="Trigger immediate Nextcloud pull and student records merge"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncingNow ? 'animate-spin text-sky-600' : ''}`} />
              {isSyncingNow ? 'Syncing...' : 'Sync Now'}
            </button>

            {/* Nextcloud Pull & Schedule Modal Button */}
            <button
              onClick={() => setIsNextcloudModalOpen(true)}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-sky-600 text-white hover:bg-sky-700 active:bg-sky-800 transition-all shadow-sm"
              title="Configure Nextcloud files and auto-sync schedule"
            >
              <DownloadCloud className="w-4 h-4" />
              Pull from Nextcloud
            </button>

            {/* Server Storage Indicator */}
            <div className="flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-medium bg-slate-50 border-slate-200 text-slate-700">
              <Server className="w-3.5 h-3.5 text-indigo-600" />
              <span>
                {serverSaveStatus === 'saving' ? (
                  <span className="text-amber-600 flex items-center gap-1">
                    <RefreshCw className="w-3 h-3 animate-spin" /> Saving to Server Storage...
                  </span>
                ) : serverSaveStatus === 'saved' ? (
                  <span className="text-emerald-700 flex items-center gap-1">
                    <CheckCircle2 className="w-3 h-3 text-emerald-600" /> Server Storage: Saved Safe on Disk
                  </span>
                ) : (
                  <span>Server Storage Ready</span>
                )}
              </span>
            </div>

            {/* Manual Save to Server Button */}
            {mergedRecords.length > 0 && (
              <button
                onClick={handleManualSaveToServer}
                disabled={isProcessing || serverSaveStatus === 'saving'}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition-colors"
                title="Force save consolidated records and histories to server storage"
              >
                <HardDrive className="w-3.5 h-3.5" />
                Save to Server Disk
              </button>
            )}

            {/* Clear All */}
            {gradeFiles.length > 0 && (
              <button
                onClick={clearAllData}
                disabled={isProcessing}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-rose-600 hover:bg-rose-50 border border-rose-200 transition-colors"
                title="Clear all loaded workbooks and records"
              >
                <Trash2 className="w-3.5 h-3.5" />
                Clear
              </button>
            )}

            {/* Select Excel Files Button */}
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept=".xlsx,.xls,.csv,.ods,.xlsm,.xlsb"
              className="hidden"
              onChange={handleFileUpload}
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessing}
              className="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 active:bg-indigo-800 transition-all shadow-sm"
            >
              <Upload className="w-4 h-4" />
              Select Excel Workbooks
            </button>
          </div>
        </div>

        {/* Processing or Status Notice Banner */}
        {processStatusMessage && (
          <div className="mt-4 p-3 bg-indigo-50 border border-indigo-100 rounded-lg flex items-center gap-2.5 text-xs text-indigo-800">
            <RefreshCw className="w-4 h-4 animate-spin text-indigo-600 shrink-0" />
            <span className="font-medium">{processStatusMessage}</span>
          </div>
        )}

        {serverSyncMessage && !processStatusMessage && (
          <div className="mt-4 p-2.5 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-between text-xs text-slate-600">
            <div className="flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              <span>{serverSyncMessage}</span>
            </div>
            {lastServerSaveTime && (
              <span className="text-slate-400 text-[11px]">
                Last synchronized: {new Date(lastServerSaveTime).toLocaleTimeString()}
              </span>
            )}
          </div>
        )}
      </div>

      {/* Drag & Drop Upload Zone (Shown prominently when no files uploaded or as drop target) */}
      {gradeFiles.length === 0 && mergedRecords.length === 0 ? (
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className={`border-2 border-dashed rounded-xl p-12 text-center transition-all ${
            isDraggingOver 
              ? 'border-indigo-500 bg-indigo-50/70 scale-[0.99]' 
              : 'border-slate-300 hover:border-indigo-400 hover:bg-slate-50/50 bg-white'
          }`}
        >
          <div className="w-16 h-16 bg-sky-50 text-sky-600 rounded-2xl flex items-center justify-center mx-auto mb-4 border border-sky-100 shadow-sm">
            <Cloud className="w-8 h-8" />
          </div>
          <h3 className="text-lg font-bold text-slate-900 mb-1">
            Pull from Nextcloud or Select Excel Workbooks
          </h3>
          <p className="text-sm text-slate-500 max-w-lg mx-auto mb-6">
            Pull your student class grade files directly from Nextcloud WebDAV on a scheduled time interval, or upload local Excel workbooks. All division sheets (A, B, C, D, E, F, G, H...) are merged preserving exact reference table columns without extra injected columns.
          </p>

          <div className="flex flex-wrap items-center justify-center gap-3">
            <button
              type="button"
              onClick={() => setIsNextcloudModalOpen(true)}
              className="inline-flex items-center gap-2 px-6 py-3 bg-sky-600 text-white rounded-xl text-sm font-semibold hover:bg-sky-700 shadow-md transition-all hover:scale-105"
            >
              <DownloadCloud className="w-4 h-4" />
              ⚡ Pull from Nextcloud (Auto-Sync)
            </button>

            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center gap-2 px-5 py-3 bg-white text-slate-700 border border-slate-300 rounded-xl text-sm font-medium hover:bg-slate-50 shadow-sm transition-all"
            >
              <Upload className="w-4 h-4 text-indigo-600" />
              Upload Local Excel Files
            </button>
          </div>


          <div className="mt-8 pt-6 border-t border-slate-200/80 grid grid-cols-1 md:grid-cols-3 gap-4 text-left max-w-3xl mx-auto">
            <div className="p-3 bg-slate-50 rounded-lg border border-slate-100">
              <div className="text-xs font-bold text-slate-800 mb-1 flex items-center gap-1.5">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                Strict Column Fidelity
              </div>
              <p className="text-xs text-slate-500">
                Only reference columns from your Excel tables are preserved. No synthetic columns added.
              </p>
            </div>
            <div className="p-3 bg-slate-50 rounded-lg border border-slate-100">
              <div className="text-xs font-bold text-slate-800 mb-1 flex items-center gap-1.5">
                <Layers className="w-3.5 h-3.5 text-indigo-600" />
                Any Divisions & Sheets
              </div>
              <p className="text-xs text-slate-500">
                Supports unlimited division sheets per book (A, B, C, D, E, F, G, H...) and any filenames.
              </p>
            </div>
            <div className="p-3 bg-slate-50 rounded-lg border border-slate-100">
              <div className="text-xs font-bold text-slate-800 mb-1 flex items-center gap-1.5">
                <Server className="w-3.5 h-3.5 text-amber-600" />
                Permanent Server Storage
              </div>
              <p className="text-xs text-slate-500">
                Uploaded files, merged records, and student history timelines are kept safe on disk.
              </p>
            </div>
          </div>
        </div>
      ) : (
        <>
          {/* Summary Stat Cards */}
          {summary && (
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between text-slate-400 mb-1">
                  <span className="text-xs font-medium">Workbooks</span>
                  <BookOpen className="w-4 h-4 text-indigo-500" />
                </div>
                <div className="text-2xl font-bold text-slate-900">{summary.totalFiles}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">Files uploaded</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between text-slate-400 mb-1">
                  <span className="text-xs font-medium">Divisions / Sheets</span>
                  <Layers className="w-4 h-4 text-blue-500" />
                </div>
                <div className="text-2xl font-bold text-slate-900">{summary.totalSheets}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">Sections consolidated</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between text-slate-400 mb-1">
                  <span className="text-xs font-medium">Total Students</span>
                  <Users className="w-4 h-4 text-violet-500" />
                </div>
                <div className="text-2xl font-bold text-slate-900">{summary.totalStudents}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">In master dataset</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between text-slate-400 mb-1">
                  <span className="text-xs font-medium">Active Enrolled</span>
                  <UserCheck className="w-4 h-4 text-emerald-500" />
                </div>
                <div className="text-2xl font-bold text-emerald-600">{summary.activeStudents}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">Current students</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between text-slate-400 mb-1">
                  <span className="text-xs font-medium">Past Students / Left</span>
                  <UserX className="w-4 h-4 text-amber-500" />
                </div>
                <div className="text-2xl font-bold text-amber-600">{summary.pastStudents}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">Left, TC or Graduated</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center justify-between text-slate-400 mb-1">
                  <span className="text-xs font-medium">Reference Columns</span>
                  <FileText className="w-4 h-4 text-slate-500" />
                </div>
                <div className="text-2xl font-bold text-slate-900">{referenceColumns.length}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">Preserved headers</div>
              </div>
            </div>
          )}

          {/* Loaded Files & Division Sheets Inspector Accordion */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Layers className="w-4 h-4 text-indigo-600" />
                <h2 className="text-sm font-bold text-slate-800">
                  Uploaded Workbooks & Class Divisions ({gradeFiles.length} files)
                </h2>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 px-3 py-1 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-700 hover:bg-slate-100 transition-colors"
                >
                  <Plus className="w-3.5 h-3.5" />
                  Add More Workbooks
                </button>
              </div>
            </div>

            <div className="divide-y divide-slate-100 max-h-72 overflow-y-auto">
              {gradeFiles.map((file, fIdx) => (
                <div key={file.id} className="p-3.5 hover:bg-slate-50/50 transition-colors">
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 mb-2">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-600 shrink-0" />
                      <span className="text-xs font-bold text-slate-800">{file.fileName}</span>
                      <span className="text-[11px] text-slate-400">
                        ({file.sheets.length} sheets, {file.sheets.reduce((acc, s) => acc + s.rowCount, 0)} total rows)
                      </span>
                    </div>

                    <div className="flex items-center gap-2">
                      <div className="flex items-center gap-1 text-xs">
                        <span className="text-slate-500 text-[11px]">Label:</span>
                        <input
                          type="text"
                          value={file.detectedGrade}
                          onChange={(e) => updateGradeFileName(file.id, e.target.value)}
                          className="px-2 py-0.5 border border-slate-300 rounded text-xs w-28 font-medium text-slate-800"
                        />
                      </div>
                      <button
                        onClick={() => removeGradeFile(file.id)}
                        className="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded"
                        title="Remove workbook"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>

                  {/* Sheets / Divisions tags */}
                  <div className="flex flex-wrap items-center gap-1.5 pl-6">
                    <span className="text-[11px] text-slate-400 mr-1">Divisions:</span>
                    {file.sheets.map((sheet, sIdx) => (
                      <div
                        key={`${file.id}-${sheet.sheetName}-${sIdx}`}
                        className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium border transition-colors ${
                          sheet.included
                            ? 'bg-indigo-50 border-indigo-200 text-indigo-800'
                            : 'bg-slate-100 border-slate-200 text-slate-400 line-through'
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={sheet.included}
                          onChange={() => toggleSheetInclusion(file.id, sheet.sheetName)}
                          className="w-3 h-3 rounded text-indigo-600 focus:ring-0 cursor-pointer"
                        />
                        <span>Sheet "{sheet.sheetName}"</span>
                        <span className="text-[10px] text-slate-500">→ Div {sheet.detectedDivision}</span>
                        <span className="text-[10px] text-slate-400">({sheet.rowCount} rows)</span>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Reference Column Inspector Bar */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 mb-3">
              <div>
                <div className="flex items-center gap-2">
                  <ShieldCheck className="w-4 h-4 text-emerald-600" />
                  <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider">
                    Reference Columns Preserved ({referenceColumns.length})
                  </h3>
                </div>
                <p className="text-[11px] text-slate-500 mt-0.5">
                  The consolidated master table and exports strictly preserve only these columns from your original reference Excel sheets:
                </p>
              </div>

              {/* Multi-Format Export Buttons */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => exportUnifiedRecordsToExcel(filteredRecords, referenceColumns)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-medium hover:bg-emerald-700 transition-colors shadow-sm"
                  title="Export to Excel (.xlsx)"
                >
                  <Download className="w-3.5 h-3.5" />
                  Export Excel (.xlsx)
                </button>
                <button
                  onClick={() => exportUnifiedRecordsToCsv(filteredRecords, referenceColumns)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 text-white rounded-lg text-xs font-medium hover:bg-slate-900 transition-colors"
                  title="Export to CSV"
                >
                  <Download className="w-3.5 h-3.5" />
                  Export CSV
                </button>
                <button
                  onClick={() => exportUnifiedRecordsToJson(filteredRecords, referenceColumns)}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 rounded-lg text-xs font-medium hover:bg-slate-100 transition-colors"
                  title="Export JSON"
                >
                  <Download className="w-3.5 h-3.5" />
                  JSON
                </button>
                <button
                  onClick={() => setIsSqlModalOpen(true)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 border border-indigo-200 text-indigo-700 rounded-lg text-xs font-medium hover:bg-indigo-100 transition-colors"
                  title="Generate Supabase SQL Schema"
                >
                  <Database className="w-3.5 h-3.5 text-indigo-600" />
                  Supabase Schema
                </button>
              </div>
            </div>

            {/* Render detected reference columns tags */}
            <div className="flex flex-wrap gap-1.5">
              {referenceColumns.map((col, idx) => (
                <span
                  key={`${col}-${idx}`}
                  className="inline-flex items-center gap-1 px-2.5 py-1 bg-slate-100 border border-slate-200 text-slate-800 rounded-md text-xs font-medium"
                >
                  <span className="text-slate-400 text-[10px] font-mono">#{idx + 1}</span>
                  {col}
                </span>
              ))}
            </div>
          </div>

          {/* Master Consolidated Table Section */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            {/* Table Control Header */}
            <div className="p-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3 bg-slate-50/70">
              {/* Cohort Tabs: All Students vs Active vs Past Students */}
              <div className="flex items-center gap-1 bg-slate-200/80 p-1 rounded-lg">
                <button
                  onClick={() => { setStudentCohortTab('ALL'); setCurrentPage(1); }}
                  className={`px-3 py-1.5 rounded-md text-xs font-bold transition-all ${
                    studentCohortTab === 'ALL'
                      ? 'bg-white text-slate-900 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  All Cohorts ({mergedRecords.length})
                </button>
                <button
                  onClick={() => { setStudentCohortTab('ACTIVE'); setCurrentPage(1); }}
                  className={`inline-flex items-center gap-1 px-3 py-1.5 rounded-md text-xs font-bold transition-all ${
                    studentCohortTab === 'ACTIVE'
                      ? 'bg-white text-emerald-700 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <UserCheck className="w-3.5 h-3.5 text-emerald-600" />
                  Active Enrolled ({summary?.activeStudents ?? 0})
                </button>
                <button
                  onClick={() => { setStudentCohortTab('PAST'); setCurrentPage(1); }}
                  className={`inline-flex items-center gap-1 px-3 py-1.5 rounded-md text-xs font-bold transition-all ${
                    studentCohortTab === 'PAST'
                      ? 'bg-white text-amber-700 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <UserX className="w-3.5 h-3.5 text-amber-600" />
                  Past / Left School Archive ({summary?.pastStudents ?? 0})
                </button>
              </div>

              {/* Filters & Search */}
              <div className="flex flex-wrap items-center gap-2">
                {/* Search Box */}
                <div className="relative min-w-[200px]">
                  <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                    placeholder="Search across columns..."
                    className="w-full pl-8 pr-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  />
                  {searchQuery && (
                    <button
                      onClick={() => setSearchQuery('')}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  )}
                </div>

                {/* Grade Filter */}
                <select
                  value={selectedGradeFilter}
                  onChange={(e) => { setSelectedGradeFilter(e.target.value); setCurrentPage(1); }}
                  className="px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-700"
                >
                  <option value="ALL">All Grades</option>
                  {availableGrades.map((g, idx) => (
                    <option key={`${g}-${idx}`} value={g}>{g}</option>
                  ))}
                </select>

                {/* Division Filter */}
                <select
                  value={selectedDivisionFilter}
                  onChange={(e) => { setSelectedDivisionFilter(e.target.value); setCurrentPage(1); }}
                  className="px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-700"
                >
                  <option value="ALL">All Divisions</option>
                  {availableDivisions.map((d, idx) => (
                    <option key={`${d}-${idx}`} value={d}>Div {d}</option>
                  ))}
                </select>

                {/* Page Size */}
                <select
                  value={pageSize}
                  onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(1); }}
                  className="px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-700"
                >
                  <option value="15">15 / page</option>
                  <option value="25">25 / page</option>
                  <option value="50">50 / page</option>
                  <option value="100">100 / page</option>
                </select>
              </div>
            </div>

            {/* Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-100/80 border-b border-slate-200 text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                    <th className="py-2.5 px-3 w-12 text-center text-slate-400">#</th>
                    {referenceColumns.map((col, cIdx) => (
                      <th
                        key={`${col}-${cIdx}`}
                        onClick={() => handleSort(col)}
                        className="py-2.5 px-3 cursor-pointer hover:bg-slate-200/60 transition-colors select-none"
                      >
                        <div className="flex items-center gap-1.5">
                          <span>{col}</span>
                          <ArrowUpDown className={`w-3 h-3 ${sortField === col ? 'text-indigo-600' : 'text-slate-400'}`} />
                        </div>
                      </th>
                    ))}
                    <th className="py-2.5 px-3 text-right">Student History</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs">
                  {paginatedRecords.length === 0 ? (
                    <tr>
                      <td colSpan={referenceColumns.length + 2} className="py-12 text-center text-slate-400">
                        <Users className="w-8 h-8 mx-auto mb-2 text-slate-300" />
                        <p className="font-medium text-slate-600">No records matching current filter</p>
                        <p className="text-xs text-slate-400 mt-0.5">Try clearing your search query or switching cohort tabs</p>
                      </td>
                    </tr>
                  ) : (
                    paginatedRecords.map((record, rIdx) => {
                      const rowNum = (currentPage - 1) * pageSize + rIdx + 1;
                      const isPast = isPastStudentStatus(record.academic_status || record['Status']);
                      const studentHistCount = studentHistories.filter(h => h.studentId === record.id || h.admissionNo === record.admission_number).length;

                      return (
                        <tr 
                          key={record.id || `rec-${rIdx}`}
                          className={`hover:bg-indigo-50/30 transition-colors ${
                            isPast ? 'bg-amber-50/20' : ''
                          }`}
                        >
                          <td className="py-2.5 px-3 text-center text-slate-400 text-[11px] font-mono">
                            {rowNum}
                          </td>

                          {referenceColumns.map((col, cIdx) => {
                            const val = record[col] !== undefined ? record[col] : '';
                            const stdKey = matchStandardColumn(col);

                            // Format status column with badges
                            if (stdKey === 'academic_status' || col.toLowerCase() === 'status') {
                              const s = String(val);
                              const isInactive = isPastStudentStatus(s);
                              return (
                                <td key={`${col}-${cIdx}`} className="py-2.5 px-3 font-medium">
                                  <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[11px] font-semibold ${
                                    isInactive
                                      ? 'bg-amber-100 text-amber-800 border border-amber-200'
                                      : 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                  }`}>
                                    {s || 'Active'}
                                  </span>
                                </td>
                              );
                            }

                            // Format Student Name with bolding
                            if (stdKey === 'student_name' || col.toLowerCase().includes('name')) {
                              return (
                                <td key={`${col}-${cIdx}`} className="py-2.5 px-3 font-bold text-slate-900">
                                  {val}
                                </td>
                              );
                            }

                            // Format Admission Number with mono font
                            if (stdKey === 'admission_number' || col.toLowerCase().includes('admission') || col.toLowerCase().includes('adm')) {
                              return (
                                <td key={`${col}-${cIdx}`} className="py-2.5 px-3 font-mono text-slate-600 text-[11px]">
                                  {val}
                                </td>
                              );
                            }

                            return (
                              <td key={`${col}-${cIdx}`} className="py-2.5 px-3 text-slate-700">
                                {String(val ?? '')}
                              </td>
                            );
                          })}

                          {/* Student History Action */}
                          <td className="py-2.5 px-3 text-right">
                            <button
                              onClick={() => openStudentHistoryModal(record)}
                              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded text-[11px] font-medium transition-colors ${
                                studentHistCount > 0
                                  ? 'bg-indigo-100 text-indigo-700 hover:bg-indigo-200'
                                  : 'bg-slate-100 text-slate-600 hover:bg-indigo-50 hover:text-indigo-600'
                              }`}
                              title="View and log history timeline for this student"
                            >
                              <History className="w-3 h-3" />
                              <span>{studentHistCount > 0 ? `${studentHistCount} Record${studentHistCount > 1 ? 's' : ''}` : 'Log History'}</span>
                            </button>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Table Footer & Pagination */}
            <div className="p-4 border-t border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 text-xs text-slate-600">
              <div>
                Showing <span className="font-bold text-slate-800">
                  {filteredRecords.length > 0 ? (currentPage - 1) * pageSize + 1 : 0}
                </span> to <span className="font-bold text-slate-800">
                  {Math.min(currentPage * pageSize, filteredRecords.length)}
                </span> of <span className="font-bold text-slate-800">{filteredRecords.length}</span> students
                {filteredRecords.length !== mergedRecords.length && (
                  <span className="text-slate-400 ml-1">
                    (filtered from {mergedRecords.length} total)
                  </span>
                )}
              </div>

              {/* Pagination Controls */}
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setCurrentPage(1)}
                  disabled={currentPage === 1}
                  className="px-2 py-1 bg-white border border-slate-300 rounded text-xs font-medium disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100"
                >
                  First
                </button>
                <button
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="p-1 bg-white border border-slate-300 rounded text-xs font-medium disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <span className="px-2 py-1 font-medium text-slate-700">
                  Page {currentPage} of {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="p-1 bg-white border border-slate-300 rounded text-xs font-medium disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
                <button
                  onClick={() => setCurrentPage(totalPages)}
                  disabled={currentPage === totalPages}
                  className="px-2 py-1 bg-white border border-slate-300 rounded text-xs font-medium disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100"
                >
                  Last
                </button>
              </div>
            </div>
          </div>
        </>
      )}

      {/* Student History Modal */}
      {isHistoryModalOpen && activeHistoryStudent && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-2xl w-full p-6 my-8 animate-in fade-in duration-200">
            {/* Modal Header */}
            <div className="flex items-center justify-between pb-4 border-b border-slate-200">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg">
                  <History className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-900">
                    Student Reference & History Timeline
                  </h3>
                  <p className="text-xs text-slate-500">
                    {activeHistoryStudent.student_name || activeHistoryStudent['Student Name']} • Adm No: {activeHistoryStudent.admission_number || activeHistoryStudent['Admission No']}
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsHistoryModalOpen(false)}
                className="p-1 text-slate-400 hover:text-slate-600 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Existing Timeline */}
            <div className="py-4 space-y-4 max-h-60 overflow-y-auto pr-1">
              <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                Logged Historical Records ({studentHistories.filter(h => h.studentId === activeHistoryStudent.id || h.admissionNo === activeHistoryStudent.admission_number).length})
              </h4>

              {studentHistories.filter(h => h.studentId === activeHistoryStudent.id || h.admissionNo === activeHistoryStudent.admission_number).length === 0 ? (
                <div className="p-4 bg-slate-50 rounded-lg text-center text-xs text-slate-500">
                  No historical records logged yet for this student. Use the form below to record when the student left school, graduated, or was issued a Transfer Certificate (TC).
                </div>
              ) : (
                <div className="space-y-2.5">
                  {studentHistories
                    .filter(h => h.studentId === activeHistoryStudent.id || h.admissionNo === activeHistoryStudent.admission_number)
                    .map((item) => (
                      <div key={item.id} className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-1">
                        <div className="flex items-center justify-between">
                          <span className="font-bold text-indigo-900 flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-indigo-600"></span>
                            {item.eventType} {item.academicYear ? `(${item.academicYear})` : ''}
                          </span>
                          <span className="text-[11px] text-slate-400">{item.date}</span>
                        </div>
                        {item.reason && (
                          <div className="text-slate-700">
                            <span className="font-medium text-slate-500">Reason / Details:</span> {item.reason}
                          </div>
                        )}
                        {item.tcOrCertNumber && (
                          <div className="text-slate-700 font-mono text-[11px]">
                            <span className="font-medium font-sans text-slate-500">TC / Certificate No:</span> {item.tcOrCertNumber}
                          </div>
                        )}
                        {item.destinationSchool && (
                          <div className="text-slate-700">
                            <span className="font-medium text-slate-500">Joined / Destination:</span> {item.destinationSchool}
                          </div>
                        )}
                        {item.description && (
                          <p className="text-slate-600 italic mt-1">{item.description}</p>
                        )}
                      </div>
                    ))}
                </div>
              )}
            </div>

            {/* Add New History Form */}
            <form onSubmit={handleAddHistoryRecord} className="pt-4 border-t border-slate-200 space-y-3">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                <Plus className="w-3.5 h-3.5 text-indigo-600" />
                Add New Event or Leaving Record
              </h4>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Event Type</label>
                  <select
                    value={historyForm.eventType}
                    onChange={(e) => setHistoryForm({ ...historyForm, eventType: e.target.value as any })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs text-slate-800"
                  >
                    <option value="Left School">Left School</option>
                    <option value="Graduated">Graduated / Passed Out</option>
                    <option value="Transferred">Transferred to Other School</option>
                    <option value="TC Issued">Transfer Certificate (TC) Issued</option>
                    <option value="Promoted">Promoted</option>
                    <option value="Award Received">Award / Achievement</option>
                    <option value="Suspended">Suspended / Inactive</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Date</label>
                  <input
                    type="date"
                    value={historyForm.date}
                    onChange={(e) => setHistoryForm({ ...historyForm, date: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs text-slate-800"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Academic Year</label>
                  <input
                    type="text"
                    placeholder="e.g. 2023-2024"
                    value={historyForm.academicYear}
                    onChange={(e) => setHistoryForm({ ...historyForm, academicYear: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs text-slate-800"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">TC / Certificate No.</label>
                  <input
                    type="text"
                    placeholder="e.g. TC-2024-0892"
                    value={historyForm.tcOrCertNumber}
                    onChange={(e) => setHistoryForm({ ...historyForm, tcOrCertNumber: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs text-slate-800"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Destination School / College</label>
                  <input
                    type="text"
                    placeholder="e.g. St. Xavier High School"
                    value={historyForm.destinationSchool}
                    onChange={(e) => setHistoryForm({ ...historyForm, destinationSchool: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs text-slate-800"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">Reason / Conduct Remarks</label>
                <textarea
                  rows={2}
                  placeholder="e.g. Family relocation to different city. Conduct was exemplary."
                  value={historyForm.reason}
                  onChange={(e) => setHistoryForm({ ...historyForm, reason: e.target.value })}
                  className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs text-slate-800"
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setIsHistoryModalOpen(false)}
                  className="px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-medium text-slate-700 hover:bg-slate-50"
                >
                  Close
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-indigo-600 text-white rounded-lg text-xs font-medium hover:bg-indigo-700 shadow-sm"
                >
                  Save Record to Server
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Supabase Schema & Push Modal */}
      {isSqlModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-3xl w-full p-6 my-8">
            <div className="flex items-center justify-between pb-4 border-b border-slate-200">
              <div className="flex items-center gap-2">
                <Database className="w-5 h-5 text-indigo-600" />
                <h3 className="text-base font-bold text-slate-900">
                  Supabase PostgreSQL Schema (Exact Reference Columns)
                </h3>
              </div>
              <button
                onClick={() => setIsSqlModalOpen(false)}
                className="p-1 text-slate-400 hover:text-slate-600 rounded-lg hover:bg-slate-100"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="py-4 space-y-4">
              <div className="flex items-center gap-3">
                <label className="text-xs font-medium text-slate-700">Target Table Name:</label>
                <input
                  type="text"
                  value={targetTableName}
                  onChange={(e) => setTargetTableName(e.target.value)}
                  className="px-3 py-1.5 border border-slate-300 rounded text-xs font-mono text-slate-900 w-64"
                />
              </div>

              {/* SQL Code Block */}
              <div className="relative">
                <pre className="p-4 bg-slate-900 text-slate-100 rounded-lg text-xs font-mono overflow-x-auto max-h-72">
                  {generateSupabaseTableSql(targetTableName, referenceColumns, mergedRecords)}
                </pre>
                <button
                  onClick={handleCopySql}
                  className="absolute top-3 right-3 px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded text-xs flex items-center gap-1 transition-colors"
                >
                  {copiedSql ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedSql ? 'Copied!' : 'Copy SQL'}</span>
                </button>
              </div>

              {/* Direct Supabase Push Progress */}
              {pushStatusMessage && (
                <div className="p-3 bg-indigo-50 border border-indigo-200 rounded-lg text-xs text-indigo-900">
                  <div className="font-medium mb-1">{pushStatusMessage}</div>
                  {isPushingToSupabase && (
                    <div className="w-full bg-indigo-200 rounded-full h-1.5 mt-2">
                      <div
                        className="bg-indigo-600 h-1.5 rounded-full transition-all duration-300"
                        style={{ width: `${pushProgress}%` }}
                      ></div>
                    </div>
                  )}
                </div>
              )}
            </div>

            <div className="flex items-center justify-between pt-4 border-t border-slate-200">
              <span className="text-xs text-slate-500">
                Creates table with exact {referenceColumns.length} reference columns without artificial columns.
              </span>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setIsSqlModalOpen(false)}
                  className="px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-medium text-slate-700 hover:bg-slate-50"
                >
                  Close
                </button>
                <button
                  onClick={handlePushToSupabase}
                  disabled={isPushingToSupabase || mergedRecords.length === 0}
                  className="inline-flex items-center gap-1.5 px-4 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-medium hover:bg-emerald-700 shadow-sm disabled:opacity-50"
                >
                  <Database className="w-3.5 h-3.5" />
                  Push {mergedRecords.length} Rows to Supabase
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Nextcloud Multi-Grade Pull & Time-Interval Auto-Sync Modal */}
      <NextcloudGradeSyncModal
        isOpen={isNextcloudModalOpen}
        onClose={() => setIsNextcloudModalOpen(false)}
        nextcloudConfig={nextcloudConfig}
        onSyncComplete={handleNextcloudModalSyncComplete}
      />
    </div>
  );
};

