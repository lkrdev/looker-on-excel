import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  makeStyles,
  tokens,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
} from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";
import { Header, NavigationMode } from "./Header";
import { AuthView } from "./AuthView";
import { ModelSelector } from "./ModelSelector";
import { PromptsDialog } from "./PromptsDialog";
import { FieldPicker } from "./FieldPicker";
import {
  FilterBuilder,
  FilterCondition,
  compileFilterToLookerExpression,
  isDateField,
} from "./FilterBuilder";
import { TableSettings } from "./TableSettings";
import { ExecutionBar } from "./ExecutionBar";
import { RefreshView } from "./RefreshView";
import { LooksBrowser } from "./LooksBrowser";
import { DashboardsBrowser } from "./DashboardsBrowser";

import { getStoredAuth, clearAuth, ensureValidToken, AuthTokens } from "../services/lookerAuth";
import {
  getModels,
  getExplore,
  getCurrentUser,
  orderFieldsDimensionsFirst,
  ModelSummary,
  ExploreDetail,
  UserSummary,
} from "../services/lookerClient";
import { resolveExcelNumberFormat } from "../services/formatMapper";
import { runQueryTask, QueryTaskController, QueryPayload } from "../services/queryTaskRunner";
import {
  writeLookerDataToWorksheet,
  expandPivotedColumns,
  resolveEffectiveColumnsForWrite,
  parseCellAddress,
  formatCellAddress,
  toggleWorksheetTreeNode,
  applyWorksheetTreeDisplayLevel,
  ColumnDefinition,
  WriteOptions,
  RowGroupDisplayLevel,
} from "../services/excelWriter";
import { buildEffectiveSortsForSubtotals } from "../services/hierarchyResolver";
import {
  getActiveSheetInfo,
  getActiveCellAddress,
  getSiblingTableBoundingBoxes,
  saveSheetMetadata,
  getAllWorkbookLookerSheets,
  StoredSheetConfig,
} from "../services/sheetMetadata";
import { clearCustomFunctionCache } from "../../functions/functions";

const useStyles = makeStyles({
  appRoot: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    width: "100%",
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: "hidden",
  },
  scrollContent: {
    flex: 1,
    overflowY: "auto",
    overflowX: "hidden",
  },
});

export const App: React.FC = () => {
  const styles = useStyles();

  // Auth & User
  const [auth, setAuth] = useState<AuthTokens | null>(() => getStoredAuth());
  const [user, setUser] = useState<UserSummary | null>(null);

  // View state: 'builder' | 'refresh' and top NavigationMode: 'explore' | 'looks' | 'dashboards'
  const [activeView, setActiveView] = useState<"builder" | "refresh">("builder");
  const [navMode, setNavMode] = useState<NavigationMode>("explore");
  const [sheetConfig, setSheetConfig] = useState<StoredSheetConfig | null>(null);
  const [allSheetTables, setAllSheetTables] = useState<StoredSheetConfig[]>([]);
  const [sourceSheetId, setSourceSheetId] = useState<string | null>(null);

  // Catalog state
  const [models, setModels] = useState<ModelSummary[]>([]);
  const [selectedModel, setSelectedModel] = useState<string>("");
  const [selectedExplore, setSelectedExplore] = useState<string>("");
  const [exploreDetail, setExploreDetail] = useState<ExploreDetail | null>(null);
  const [loadingModels, setLoadingModels] = useState(false);
  const [loadingExplore, setLoadingExplore] = useState(false);

  // Query Builder state
  const [selectedFields, setSelectedFields] = useState<string[]>([]);
  const [sorts, setSorts] = useState<string[]>([]);
  const [pivots, setPivots] = useState<string[]>([]);
  const [filters, setFilters] = useState<FilterCondition[]>([]);
  const [promptValues, setPromptValues] = useState<Record<string, string>>({});
  const [rowLimit, setRowLimit] = useState<string>("500");
  const [destination, setDestination] = useState<"active" | "new">("active");
  const [tableOptions, setTableOptions] = useState<WriteOptions>({
    useExcelTable: true,
    tableStyle: "TableStyleMedium2",
    preserveUserFormatting: true,
    autoExpandFormulas: true,
    freezeHeader: true,
    nullDisplay: "",
    includeTotals: false,
    includeRowTotals: false,
    includeSubtotals: false,
    anchorCell: "A1",
  });

  // Execution & Progress state
  const [isExecuting, setIsExecuting] = useState(false);
  const [statusText, setStatusText] = useState("");
  const [progressPercent, setProgressPercent] = useState(0);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [refreshingAll, setRefreshingAll] = useState(false);

  const activeTaskRef = useRef<QueryTaskController | null>(null);
  const lastActiveSheetIdRef = useRef<string | null>(null);
  const sourceSheetIdRef = useRef<string | null>(null);
  const preferredTableIdRef = useRef<string | null>(null);
  const skipDefaultPromptsRef = useRef<boolean>(false);
  const pendingFieldOrderRef = useRef<boolean>(false);
  const togglingTreeNodeRef = useRef<boolean>(false);

  // Validate / refresh token on taskpane mount / relaunch of Excel
  useEffect(() => {
    if (auth && auth.refreshToken) {
      ensureValidToken(auth)
        .then((valid) => {
          if (valid.accessToken !== auth.accessToken) {
            setAuth(valid);
          }
        })
        .catch((e) => console.warn("Token refresh on relaunch failed:", e));
    }
  }, []);

  // Check active sheet metadata; also hit-test which table on a multi-table sheet contains the active cell
  const checkSheetMetadata = useCallback(async (forceViewSync: boolean = false) => {
    const {
      activeSheetId,
      sourceSheetId: resolvedSourceId,
      config,
      allSheetTables: resolvedTables,
      selectedCellAddress,
      isPivotSheet,
    } = await getActiveSheetInfo(sourceSheetIdRef.current, preferredTableIdRef.current);

    const activeChanged =
      activeSheetId !== null && activeSheetId !== lastActiveSheetIdRef.current;
    const sourceChanged = resolvedSourceId !== sourceSheetIdRef.current;

    if (activeChanged) {
      preferredTableIdRef.current = null;
    }

    // In-sheet interactive expand/collapse when user clicks a `▸` / `▾` parent cell in the Group column
    if (
      !forceViewSync &&
      !activeChanged &&
      !togglingTreeNodeRef.current &&
      selectedCellAddress &&
      config?.options?.enableRowGroups &&
      config?.options?.treeNodes &&
      config.options.treeNodes.length > 0
    ) {
      const { startRow: selRow, startCol: selCol } = parseCellAddress(selectedCellAddress);
      const { startRow: tblRow, startCol: tblCol } = parseCellAddress(
        config.anchorCell || config.options.anchorCell || "A1"
      );
      if (selCol === tblCol && selRow >= tblRow + 1) {
        try {
          togglingTreeNodeRef.current = true;
          const updatedNodes = await toggleWorksheetTreeNode(
            resolvedSourceId || activeSheetId || undefined,
            config.options,
            config.rowCount,
            selRow,
            selCol
          );
          if (updatedNodes) {
            const updatedCfg: StoredSheetConfig = {
              ...config,
              options: {
                ...config.options,
                treeNodes: updatedNodes,
              },
            };
            const nextTables = await saveSheetMetadata(
              updatedCfg,
              resolvedSourceId || activeSheetId || undefined
            );
            setAllSheetTables(nextTables);
            setSheetConfig(updatedCfg);
            return;
          }
        } catch (err) {
          console.warn("Tree node toggle skipped:", err);
        } finally {
          togglingTreeNodeRef.current = false;
        }
      }
    }

    setAllSheetTables(resolvedTables || []);
    if (config) {
      setSheetConfig(config);
      preferredTableIdRef.current = config.tableId || null;
      if (config.queryPayload?.model && config.queryPayload?.view) {
        setSelectedModel(config.queryPayload.model);
        setSelectedExplore(config.queryPayload.view);
      }
    }

    if (forceViewSync || activeChanged || lastActiveSheetIdRef.current === null) {
      lastActiveSheetIdRef.current = activeSheetId;
      sourceSheetIdRef.current = resolvedSourceId;
      setSourceSheetId(resolvedSourceId);
      setSheetConfig(config);

      if (forceViewSync || sourceChanged || !isPivotSheet) {
        setActiveView(config ? "refresh" : "builder");
      }
    }
  }, []);

  // Listen for worksheet switches (both selection changes and worksheet tab activation)
  useEffect(() => {
    checkSheetMetadata(true);

    const onSheetOrSelectionChange = () => {
      checkSheetMetadata(false);
    };

    if (window.Office && Office.context && Office.context.document) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSheetOrSelectionChange
      );
    }

    let activationHandlerResult: OfficeExtension.EventHandlerResult<Excel.WorksheetActivatedEventArgs> | null =
      null;
    if (typeof Excel !== "undefined" && typeof Excel.run === "function") {
      Excel.run(async (context) => {
        activationHandlerResult = context.workbook.worksheets.onActivated.add(async () => {
          onSheetOrSelectionChange();
        });
        await context.sync();
      }).catch(() => {});
    }

    return () => {
      if (window.Office && Office.context && Office.context.document) {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          onSheetOrSelectionChange
        );
      }
      if (activationHandlerResult && typeof Excel !== "undefined") {
        Excel.run(activationHandlerResult.context, async (context) => {
          activationHandlerResult?.remove();
          await context.sync();
        }).catch(() => {});
      }
    };
  }, [checkSheetMetadata]);

  // Load User profile & Catalog when authenticated
  useEffect(() => {
    if (!auth) return;

    getCurrentUser(auth.baseUrl, auth.accessToken)
      .then(setUser)
      .catch((e) => console.warn("Could not fetch user:", e));

    setLoadingModels(true);
    getModels(auth.baseUrl, auth.accessToken)
      .then((data) => {
        setModels(data);
        // Default to first model and explore if available and not already selected
        const defaultModel = data[0];
        if (defaultModel) {
          setSelectedModel((prev) => prev || defaultModel.name);
          const defaultExp = defaultModel.explores?.[0];
          if (defaultExp) {
            setSelectedExplore((prev) => prev || defaultExp.name);
          }
        }
      })
      .catch((e) => setErrorMessage(`Failed to load models: ${e.message}`))
      .finally(() => setLoadingModels(false));
  }, [auth]);

  // Load Explore fields when selected
  useEffect(() => {
    if (!auth || !selectedModel || !selectedExplore) return;

    setLoadingExplore(true);
    setErrorMessage(null);

    getExplore(auth.baseUrl, auth.accessToken, selectedModel, selectedExplore)
      .then((data) => {
        setExploreDetail(data);

        if (pendingFieldOrderRef.current) {
          pendingFieldOrderRef.current = false;
          const allDefs = [...data.fields.dimensions, ...data.fields.measures];
          setSelectedFields((prev) =>
            orderFieldsDimensionsFirst(prev, allDefs, pivots)
          );
        }

        // Pre-fill defaults for any required always_filter unless hydrating from handleEditQuery
        if (skipDefaultPromptsRef.current) {
          skipDefaultPromptsRef.current = false;
        } else if (data.always_filter && data.always_filter.length > 0) {
          const initialPrompts: Record<string, string> = {};
          data.always_filter.forEach((af) => {
            if (af.values && af.values.length > 0) {
              initialPrompts[af.field] = af.values.join(", ");
            }
          });
          setPromptValues(initialPrompts);
        }
      })
      .catch((e) => setErrorMessage(`Failed to load explore fields: ${e.message}`))
      .finally(() => setLoadingExplore(false));
  }, [auth, selectedModel, selectedExplore]);

  // Auth Handlers
  const handleAuthSuccess = (tokens: AuthTokens) => {
    setAuth(tokens);
    setErrorMessage(null);
  };

  const handleDisconnect = () => {
    clearAuth();
    setAuth(null);
    setUser(null);
    setModels([]);
    setExploreDetail(null);
    setSelectedFields([]);
    setSorts([]);
    setPivots([]);
    setFilters([]);
    setPromptValues({});
  };

  // Field selection, sort, and pivot handlers
  const handleToggleField = (fieldName: string) => {
    setSelectedFields((prev) => {
      const exists = prev.includes(fieldName);
      if (exists) {
        setSorts((s) => s.filter((x) => x !== fieldName && !x.startsWith(`${fieldName} `)));
        setPivots((p) => p.filter((x) => x !== fieldName));
        return prev.filter((f) => f !== fieldName);
      }
      // Keep Dimensions before Measures by default when adding a new Dimension
      const isDim = exploreDetail?.fields.dimensions.some((d) => d.name === fieldName);
      if (isDim && exploreDetail) {
        const measureSet = new Set(exploreDetail.fields.measures.map((m) => m.name));
        const firstMeasureIdx = prev.findIndex((f) => measureSet.has(f));
        if (firstMeasureIdx !== -1) {
          const next = [...prev];
          next.splice(firstMeasureIdx, 0, fieldName);
          return next;
        }
      }
      return [...prev, fieldName];
    });
  };

  const handleClearSelected = () => {
    setSelectedFields([]);
    setSorts([]);
    setPivots([]);
  };

  const handleToggleSort = (fieldName: string) => {
    setSorts((prev) => {
      const existingIndex = prev.findIndex((s) => s === fieldName || s.startsWith(`${fieldName} `));
      if (existingIndex === -1) {
        // Not sorted -> sort ASC
        return [...prev, `${fieldName} asc`];
      }
      const current = prev[existingIndex];
      if (current === fieldName || current === `${fieldName} asc`) {
        // ASC -> DESC
        const updated = [...prev];
        updated[existingIndex] = `${fieldName} desc`;
        return updated;
      }
      // DESC -> remove sort
      return prev.filter((_, idx) => idx !== existingIndex);
    });
  };

  const handleTogglePivot = (fieldName: string) => {
    setPivots((prev) =>
      prev.includes(fieldName) ? prev.filter((p) => p !== fieldName) : [...prev, fieldName]
    );
  };

  // Hydrate query builder from stored sheet metadata (Edit Query workflow)
  const handleEditQuery = (config: StoredSheetConfig) => {
    if (
      config.queryPayload.model !== selectedModel ||
      config.queryPayload.view !== selectedExplore
    ) {
      skipDefaultPromptsRef.current = true;
    }

    setSelectedModel(config.queryPayload.model);
    setSelectedExplore(config.queryPayload.view);
    setSelectedFields(config.queryPayload.fields || []);
    setSorts(config.queryPayload.sorts || []);
    setPivots(config.queryPayload.pivots || []);
    setRowLimit(config.queryPayload.limit || "5000");
    if (config.options) {
      setTableOptions({
        ...config.options,
        anchorCell: config.anchorCell || config.options.anchorCell || "A1",
        tableId: config.tableId || config.options.tableId,
      });
      setDestination("active");
    }

    if (config.filters || config.promptValues) {
      setPromptValues(config.promptValues || {});
      setFilters(config.filters || []);
    } else if (config.queryPayload.filters) {
      const restoredPrompts: Record<string, string> = {};
      const restoredFilters: FilterCondition[] = [];
      const promptFieldNames = new Set<string>([
        ...(exploreDetail?.always_filter?.map((af) => af.field) || []),
        ...(exploreDetail?.fields.parameters?.map((p) => p.name) || []),
      ]);

      Object.entries(config.queryPayload.filters).forEach(([field, expr], idx) => {
        if (promptFieldNames.has(field)) {
          restoredPrompts[field] = expr;
        } else {
          restoredFilters.push({
            id: `${Date.now()}_${idx}`,
            field,
            operator: "matches_advanced",
            value: expr,
          });
        }
      });
      setPromptValues(restoredPrompts);
      setFilters(restoredFilters);
    } else {
      setPromptValues({});
      setFilters([]);
    }

    setActiveView("builder");
  };

  // Start fresh query from scratch, defaulting destination to new sheet to protect existing tables
  const handleStartNewQuery = () => {
    setSheetConfig(null);
    sourceSheetIdRef.current = null;
    preferredTableIdRef.current = null;
    setSourceSheetId(null);
    setSelectedFields([]);
    setSorts([]);
    setPivots([]);
    setFilters([]);
    setPromptValues({});
    setTableOptions((prev) => ({ ...prev, anchorCell: "A1", tableId: undefined }));
    setDestination("new");
    setActiveView("builder");
  };

  // Add another table onto the current worksheet at the user's selected cell (or next available column)
  const handleAddTableToSheet = async () => {
    const selectedAddr = await getActiveCellAddress();
    let suggestedAnchor = selectedAddr || "H1";
    const existingAnchors = new Set(
      allSheetTables.map((t) => (t.anchorCell || t.options?.anchorCell || "A1").toUpperCase())
    );
    if (existingAnchors.has(suggestedAnchor.toUpperCase())) {
      const maxEndCol = allSheetTables.reduce((maxC, t) => Math.max(maxC, t.endCol ?? 4), 0);
      suggestedAnchor = formatCellAddress(0, maxEndCol + 2);
    }

    setSheetConfig(null);
    preferredTableIdRef.current = null;
    setSelectedFields([]);
    setSorts([]);
    setPivots([]);
    setFilters([]);
    setPromptValues({});
    setDestination("active");
    setTableOptions((prev) => ({
      ...prev,
      anchorCell: suggestedAnchor,
      tableId: `Looker_${suggestedAnchor}`,
    }));
    setActiveView("builder");
  };

  // Filter handlers
  const handleAddFilter = () => {
    const newId = Date.now().toString();
    const allFields = exploreDetail
      ? [...exploreDetail.fields.dimensions, ...exploreDetail.fields.measures]
      : [];
    const firstFieldDef = allFields[0];
    const firstField = firstFieldDef?.name || "";

    let defaultOp = "is";
    let defaultValue = "";
    let defaultUnit: string | undefined = undefined;

    if (isDateField(firstFieldDef)) {
      defaultOp = "is_in_the_last";
      defaultValue = "7";
      defaultUnit = "days";
    }

    setFilters((prev) => [
      ...prev,
      { id: newId, field: firstField, operator: defaultOp, value: defaultValue, unit: defaultUnit },
    ]);
  };

  const handleUpdateFilter = (id: string, updates: Partial<FilterCondition>) => {
    setFilters((prev) => prev.map((f) => (f.id === id ? { ...f, ...updates } : f)));
  };

  const handleRemoveFilter = (id: string) => {
    setFilters((prev) => prev.filter((f) => f.id !== id));
  };

  // Execute Import
  const handleExecuteImport = async () => {
    if (!auth || !exploreDetail || selectedFields.length === 0) return;

    setIsExecuting(true);
    setStatusText("Verifying session and submitting query...");
    setProgressPercent(0);
    setErrorMessage(null);

    try {
      // 0. Ensure fresh token before executing query
      const validAuth = await ensureValidToken(auth);
      if (validAuth.accessToken !== auth.accessToken) {
        setAuth(validAuth);
      }

      // 1. Prepare fields lookup & compile filters
      const allAvailableFields = [
        ...exploreDetail.fields.dimensions,
        ...exploreDetail.fields.measures,
      ];

      // Only include prompt / parameter entries that have actual non-empty values
      const compiledFilters: Record<string, string> = {};
      Object.entries(promptValues).forEach(([k, v]) => {
        if (v !== undefined && v !== null && v.trim() !== "") {
          compiledFilters[k] = v.trim();
        }
      });

      // Ensure any always_filter has its default applied if not explicitly set in promptValues
      if (exploreDetail.always_filter) {
        exploreDetail.always_filter.forEach((af) => {
          if (!compiledFilters[af.field] && af.values && af.values.length > 0) {
            compiledFilters[af.field] = af.values.join(", ");
          }
        });
      }

      filters.forEach((f) => {
        if (f.field) {
          const def = allAvailableFields.find((fld) => fld.name === f.field);
          const expr = compileFilterToLookerExpression(f, def);
          if (expr !== null && expr !== undefined && expr !== "") {
            compiledFilters[f.field] = compiledFilters[f.field]
              ? `${compiledFilters[f.field]}, ${expr}`
              : expr;
          }
        }
      });

      const nonPivotDims = selectedFields.filter((fName) => {
        const def = allAvailableFields.find((f) => f.name === fName);
        return def?.category === "dimension" && !pivots.includes(fName);
      });

      const orderedNonPivotDims =
        tableOptions.enableRowGroups &&
        tableOptions.rowGroupFields &&
        tableOptions.rowGroupFields.length >= 2
          ? [
              ...tableOptions.rowGroupFields.filter((f) => nonPivotDims.includes(f)),
              ...nonPivotDims.filter((f) => !tableOptions.rowGroupFields!.includes(f)),
            ]
          : nonPivotDims;

      const hasSelectedMeasures = selectedFields.some((fName) => {
        const def = allAvailableFields.find((f) => f.name === fName);
        return def?.category === "measure";
      });

      const wantSubtotals =
        hasSelectedMeasures &&
        (tableOptions.includeSubtotals ||
          (tableOptions.enableRowGroups && tableOptions.showGroupSubtotals !== false)) &&
        orderedNonPivotDims.length >= 2;

      const subtotalsList = wantSubtotals ? orderedNonPivotDims.slice(0, -1) : null;
      const effectiveSorts = subtotalsList
        ? buildEffectiveSortsForSubtotals(sorts, subtotalsList)
        : sorts.length > 0
        ? sorts
        : [];

      const queryPayload: QueryPayload = {
        model: selectedModel,
        view: selectedExplore,
        fields: selectedFields,
        filters: Object.keys(compiledFilters).length > 0 ? compiledFilters : null,
        sorts: effectiveSorts,
        limit: rowLimit,
        pivots: pivots.length > 0 ? pivots : null,
        total: tableOptions.includeTotals ? true : undefined,
        row_total:
          tableOptions.includeRowTotals && pivots.length > 0 ? "right" : undefined,
        subtotals: subtotalsList,
      };

      // 2. Prepare column metadata with number formats
      const columns: ColumnDefinition[] = selectedFields.map((fieldKey) => {
        const def = allAvailableFields.find((f) => f.name === fieldKey);
        return {
          fieldKey,
          label: def?.label || fieldKey,
          type: def?.type,
          excelFormat: def ? resolveExcelNumberFormat(def) : undefined,
          category: def?.category,
        };
      });

      // 3. Launch Async Query Task
      const controller = await runQueryTask(
        validAuth.baseUrl,
        validAuth.accessToken,
        queryPayload
      );
      activeTaskRef.current = controller;

      const bundle = await controller.waitForBundle((status, elapsed) => {
        setStatusText(`Running query in warehouse... (${elapsed}s)`);
      });
      const dataRows = bundle.rows;

      // 4. Stream to Excel
      setStatusText(`Streaming ${dataRows.length.toLocaleString()} rows into worksheet...`);
      const baseColsForExpand = resolveEffectiveColumnsForWrite(columns, {
        ...tableOptions,
        pivots: pivots.length > 0 ? pivots : undefined,
        rowGroupFields: orderedNonPivotDims,
      });
      const expandedColCount = expandPivotedColumns(
        baseColsForExpand,
        dataRows,
        pivots.length > 0 ? pivots : undefined,
        Boolean(tableOptions.includeRowTotals && pivots.length > 0),
        bundle.pivotsMeta,
        tableOptions.pivotGroupMode || "measure_first"
      ).expandedColumns.length;

      const targetSheetId =
        destination === "active" && sourceSheetId ? sourceSheetId : undefined;
      const effectiveAnchor =
        destination === "new" ? "A1" : (tableOptions.anchorCell || "A1").toUpperCase();
      const effectiveTableId =
        destination === "new"
          ? "Looker_A1"
          : tableOptions.tableId || sheetConfig?.tableId || `Looker_${effectiveAnchor}`;
      const siblingRanges =
        destination === "active"
          ? getSiblingTableBoundingBoxes(allSheetTables, effectiveTableId, effectiveAnchor)
          : [];

      const writeOptions: WriteOptions = {
        ...tableOptions,
        rowGroupFields: orderedNonPivotDims,
        anchorCell: effectiveAnchor,
        tableId: effectiveTableId,
        siblingTableRanges: siblingRanges,
        destination,
        targetSheetId,
        sheetName: exploreDetail.label || selectedExplore || "Looker Data",
        pivots: pivots.length > 0 ? pivots : undefined,
        pivotsMeta: bundle.pivotsMeta,
        prevRowCount: destination === "active" ? sheetConfig?.rowCount : undefined,
        prevColCount:
          destination === "active"
            ? sheetConfig?.options?.prevColCount ?? sheetConfig?.columns?.length
            : undefined,
        totalsData: bundle.totalsData,
        subtotalsData: bundle.subtotalsData,
      };
      const writtenSheetId = await writeLookerDataToWorksheet(
        columns,
        dataRows,
        writeOptions,
        (pct) => {
          setProgressPercent(pct);
        }
      );

      const totalWrittenDataRows =
        writeOptions.rowLevels && writeOptions.rowLevels.length > 0
          ? writeOptions.rowLevels.length
          : dataRows.length +
            (bundle.subtotalsData ? Object.values(bundle.subtotalsData).flat().length : 0) +
            (bundle.totalsData ? 1 : 0);

      // 5. Persist metadata for one-click refresh on the actual Looker data worksheet
      const config: StoredSheetConfig = {
        tableId: effectiveTableId,
        anchorCell: effectiveAnchor,
        queryPayload,
        columns,
        options: {
          ...writeOptions,
          totalsData: undefined,
          subtotalsData: undefined,
          siblingTableRanges: undefined,
          destination: "active",
          targetSheetId: undefined,
          prevRowCount: totalWrittenDataRows,
          prevColCount: expandedColCount,
        },
        lastRefreshed: new Date().toISOString(),
        rowCount: totalWrittenDataRows,
        modelLabel: models.find((m) => m.name === selectedModel)?.label,
        exploreLabel: exploreDetail.label,
        filters,
        promptValues,
      };
      await saveSheetMetadata(config, writtenSheetId || targetSheetId);
      preferredTableIdRef.current = effectiveTableId;
      if (writtenSheetId) {
        sourceSheetIdRef.current = writtenSheetId;
        setSourceSheetId(writtenSheetId);
      }
      await checkSheetMetadata(true);
    } catch (err: any) {
      setErrorMessage(err.message || "Query execution failed.");
    } finally {
      setIsExecuting(false);
      activeTaskRef.current = null;
    }
  };

  // Re-run Refresh on active (or source) Looker table (optionally with an expanded/collapsed hierarchy payload)
  const handleRefreshActive = async (
    overridePayload?: QueryPayload,
    overrideOptions?: Partial<WriteOptions>
  ) => {
    if (!auth || !sheetConfig) return;
    clearCustomFunctionCache();

    setIsExecuting(true);
    setStatusText("Submitting refresh query...");
    setProgressPercent(0);
    setErrorMessage(null);

    try {
      const validAuth = await ensureValidToken(auth);
      if (validAuth.accessToken !== auth.accessToken) {
        setAuth(validAuth);
      }

      const basePayload = overridePayload || sheetConfig.queryPayload;
      const mergedOptions: WriteOptions = {
        ...sheetConfig.options,
        ...(overrideOptions || {}),
      };
      const allAvailableFields = exploreDetail
        ? [...exploreDetail.fields.dimensions, ...exploreDetail.fields.measures]
        : [];

      const activeColumns: ColumnDefinition[] = overridePayload
        ? (basePayload.fields || []).map((fieldKey) => {
            const existingCol = sheetConfig.columns.find((c) => c.fieldKey === fieldKey);
            if (existingCol) return existingCol;
            const def = allAvailableFields.find((f) => f.name === fieldKey);
            return {
              fieldKey,
              label: def?.label || fieldKey,
              type: def?.type,
              excelFormat: def ? resolveExcelNumberFormat(def) : undefined,
              category: def?.category,
            };
          })
        : sheetConfig.columns;

      const activePivots = basePayload.pivots || undefined;
      const nonPivotDims = (basePayload.fields || []).filter((fName) => {
        const col = activeColumns.find((c) => c.fieldKey === fName);
        const def = allAvailableFields.find((f) => f.name === fName);
        const isDim = (col?.category || def?.category || "dimension") === "dimension";
        return isDim && !(activePivots || []).includes(fName);
      });

      const hasSelectedMeasures = (basePayload.fields || []).some((fName) => {
        const col = activeColumns.find((c) => c.fieldKey === fName);
        const def = allAvailableFields.find((f) => f.name === fName);
        return (col?.category || def?.category) === "measure";
      });

      const wantSubtotals =
        hasSelectedMeasures &&
        (mergedOptions.includeSubtotals ||
          (mergedOptions.enableRowGroups && mergedOptions.showGroupSubtotals !== false)) &&
        nonPivotDims.length >= 2;

      const subtotalsList = wantSubtotals ? nonPivotDims.slice(0, -1) : null;
      const effectiveSorts = subtotalsList
        ? buildEffectiveSortsForSubtotals(basePayload.sorts || [], subtotalsList)
        : basePayload.sorts || [];

      const effectivePayload: QueryPayload = {
        ...basePayload,
        sorts: effectiveSorts,
        subtotals: subtotalsList,
      };

      const controller = await runQueryTask(
        validAuth.baseUrl,
        validAuth.accessToken,
        effectivePayload
      );
      activeTaskRef.current = controller;

      const bundle = await controller.waitForBundle((_, elapsed) => {
        setStatusText(`Refreshing data from warehouse... (${elapsed}s)`);
      });
      const dataRows = bundle.rows;

      const baseColsForExpand = resolveEffectiveColumnsForWrite(activeColumns, {
        ...mergedOptions,
        pivots: activePivots,
        rowGroupFields: nonPivotDims,
      });
      const expandedColCount = expandPivotedColumns(
        baseColsForExpand,
        dataRows,
        activePivots,
        Boolean(effectivePayload.row_total),
        bundle.pivotsMeta,
        sheetConfig.options?.pivotGroupMode || "measure_first"
      ).expandedColumns.length;

      const effectiveAnchor = (
        sheetConfig.anchorCell ||
        mergedOptions.anchorCell ||
        "A1"
      ).toUpperCase();
      const effectiveTableId =
        sheetConfig.tableId || mergedOptions.tableId || `Looker_${effectiveAnchor}`;
      const siblingRanges = getSiblingTableBoundingBoxes(
        allSheetTables,
        effectiveTableId,
        effectiveAnchor
      );

      setStatusText(`Writing ${dataRows.length.toLocaleString()} rows...`);
      const writeOpts: WriteOptions = {
        ...mergedOptions,
        rowGroupFields: nonPivotDims,
        anchorCell: effectiveAnchor,
        tableId: effectiveTableId,
        siblingTableRanges: siblingRanges,
        destination: "active",
        targetSheetId: sourceSheetId || undefined,
        pivots: activePivots,
        pivotsMeta: bundle.pivotsMeta,
        prevRowCount: sheetConfig.rowCount,
        prevColCount: sheetConfig.options?.prevColCount ?? sheetConfig.columns.length,
        totalsData: bundle.totalsData,
        subtotalsData: bundle.subtotalsData,
      };
      const writtenSheetId = await writeLookerDataToWorksheet(
        activeColumns,
        dataRows,
        writeOpts,
        (pct) => setProgressPercent(pct)
      );

      const totalWrittenDataRows =
        writeOpts.rowLevels && writeOpts.rowLevels.length > 0
          ? writeOpts.rowLevels.length
          : dataRows.length +
            (bundle.subtotalsData ? Object.values(bundle.subtotalsData).flat().length : 0) +
            (bundle.totalsData ? 1 : 0);

      const updatedConfig: StoredSheetConfig = {
        ...sheetConfig,
        tableId: effectiveTableId,
        anchorCell: effectiveAnchor,
        queryPayload: effectivePayload,
        columns: activeColumns,
        options: {
          ...writeOpts,
          totalsData: undefined,
          subtotalsData: undefined,
          siblingTableRanges: undefined,
          destination: "active",
          targetSheetId: undefined,
          prevRowCount: totalWrittenDataRows,
          prevColCount: expandedColCount,
        },
        lastRefreshed: new Date().toISOString(),
        rowCount: totalWrittenDataRows,
      };
      const updatedTables = await saveSheetMetadata(
        updatedConfig,
        writtenSheetId || sourceSheetId || undefined
      );
      setAllSheetTables(updatedTables);
      setSheetConfig(updatedConfig);
    } catch (err: any) {
      setErrorMessage(err.message || "Failed to refresh data.");
    } finally {
      setIsExecuting(false);
      activeTaskRef.current = null;
    }
  };

  // Global "Refresh All" across workbook (including all tables on multi-table sheets)
  const handleRefreshAll = async () => {
    if (!auth) return;
    clearCustomFunctionCache();
    setRefreshingAll(true);
    setErrorMessage(null);
    try {
      const validAuth = await ensureValidToken(auth);
      if (validAuth.accessToken !== auth.accessToken) {
        setAuth(validAuth);
      }

      const sheets = await getAllWorkbookLookerSheets();
      for (const s of sheets) {
        const controller = await runQueryTask(
          validAuth.baseUrl,
          validAuth.accessToken,
          s.config.queryPayload
        );
        const bundle = await controller.waitForBundle(() => {});
        const dataRows = bundle.rows;
        const sPivots = s.config.queryPayload.pivots || undefined;
        const expandedColCount = expandPivotedColumns(
          s.config.columns,
          dataRows,
          sPivots,
          Boolean(s.config.queryPayload.row_total),
          bundle.pivotsMeta,
          s.config.options?.pivotGroupMode || "measure_first"
        ).expandedColumns.length;

        const effectiveAnchor = (
          s.config.anchorCell ||
          s.config.options?.anchorCell ||
          "A1"
        ).toUpperCase();
        const effectiveTableId =
          s.config.tableId || s.config.options?.tableId || `Looker_${effectiveAnchor}`;
        const siblingRanges = getSiblingTableBoundingBoxes(
          s.allTables || [],
          effectiveTableId,
          effectiveAnchor
        );

        await writeLookerDataToWorksheet(
          s.config.columns,
          dataRows,
          {
            ...s.config.options,
            anchorCell: effectiveAnchor,
            tableId: effectiveTableId,
            siblingTableRanges: siblingRanges,
            destination: "active",
            targetSheetId: s.sheetId,
            pivots: sPivots,
            pivotsMeta: bundle.pivotsMeta,
            prevRowCount: s.config.rowCount,
            prevColCount: s.config.options?.prevColCount ?? s.config.columns.length,
            totalsData: bundle.totalsData,
            subtotalsData: bundle.subtotalsData,
          },
          () => {}
        );

        const totalWrittenDataRows =
          dataRows.length +
          (bundle.subtotalsData ? Object.values(bundle.subtotalsData).flat().length : 0) +
          (bundle.totalsData ? 1 : 0);

        await saveSheetMetadata(
          {
            ...s.config,
            tableId: effectiveTableId,
            anchorCell: effectiveAnchor,
            options: {
              ...s.config.options,
              anchorCell: effectiveAnchor,
              tableId: effectiveTableId,
              destination: "active",
              prevRowCount: totalWrittenDataRows,
              prevColCount: expandedColCount,
            },
            lastRefreshed: new Date().toISOString(),
            rowCount: totalWrittenDataRows,
          },
          s.sheetId
        );
      }
      await checkSheetMetadata(true);
    } catch (e: any) {
      setErrorMessage(`Refresh All failed: ${e.message}`);
    } finally {
      setRefreshingAll(false);
    }
  };

  const handleCancelQuery = async () => {
    if (activeTaskRef.current) {
      await activeTaskRef.current.cancel();
      setIsExecuting(false);
      setStatusText("Query canceled.");
    }
  };

  // Load a Saved Look or Dashboard Tile query into the Explore Builder
  const handleOpenPayloadInExplore = (payload: QueryPayload) => {
    skipDefaultPromptsRef.current = true;
    const isSameExplore =
      exploreDetail &&
      payload.model === selectedModel &&
      payload.view === selectedExplore;
    const initialOrderedFields =
      isSameExplore && exploreDetail
        ? orderFieldsDimensionsFirst(
            payload.fields || [],
            [...exploreDetail.fields.dimensions, ...exploreDetail.fields.measures],
            payload.pivots
          )
        : payload.fields || [];
    if (!isSameExplore) {
      pendingFieldOrderRef.current = true;
    }

    setSelectedModel(payload.model);
    setSelectedExplore(payload.view);
    setSelectedFields(initialOrderedFields);
    setSorts(payload.sorts || []);
    setPivots(payload.pivots || []);
    setRowLimit(payload.limit || "500");
    setTableOptions((prev) => ({
      ...prev,
      includeTotals: Boolean(payload.total),
      includeRowTotals: Boolean(payload.row_total),
      includeSubtotals: Boolean(payload.subtotals && payload.subtotals.length > 0),
    }));

    const restoredFilters: FilterCondition[] = [];
    if (payload.filters) {
      Object.entries(payload.filters).forEach(([field, expr], idx) => {
        restoredFilters.push({
          id: `${Date.now()}_${idx}`,
          field,
          operator: "matches_advanced",
          value: expr,
        });
      });
    }
    setFilters(restoredFilters);
    setPromptValues({});
    setNavMode("explore");
    setActiveView("builder");
  };

  // Directly import a Saved Look or Dashboard Tile QueryPayload into Excel
  const handleImportExternalPayload = async (
    payload: QueryPayload,
    sheetTitle: string,
    forceNewSheet: boolean = false
  ): Promise<void> => {
    if (!auth) return;
    setIsExecuting(true);
    setStatusText(`Running "${sheetTitle}" in warehouse...`);
    setProgressPercent(0);
    setErrorMessage(null);

    try {
      const validAuth = await ensureValidToken(auth);
      if (validAuth.accessToken !== auth.accessToken) {
        setAuth(validAuth);
      }

      // Fetch explore metadata to resolve friendly labels and number formats
      let expMeta: ExploreDetail | null = null;
      try {
        expMeta = await getExplore(
          validAuth.baseUrl,
          validAuth.accessToken,
          payload.model,
          payload.view
        );
      } catch {
        // Fallback to raw field names if explore metadata fails
      }

      const allAvailableFields = expMeta
        ? [...expMeta.fields.dimensions, ...expMeta.fields.measures]
        : [];

      // Normalize field order so Dimensions always precede Measures (Looker API returns raw click order)
      const orderedFieldKeys = orderFieldsDimensionsFirst(
        payload.fields || [],
        allAvailableFields,
        payload.pivots
      );
      const orderedPayload: QueryPayload = {
        ...payload,
        fields: orderedFieldKeys,
      };

      const columns: ColumnDefinition[] = orderedFieldKeys.map((fieldKey) => {
        const def = allAvailableFields.find((f) => f.name === fieldKey);
        return {
          fieldKey,
          label: def?.label || fieldKey,
          type: def?.type,
          excelFormat: def ? resolveExcelNumberFormat(def) : undefined,
          category: def?.category,
        };
      });

      const controller = await runQueryTask(
        validAuth.baseUrl,
        validAuth.accessToken,
        orderedPayload
      );
      activeTaskRef.current = controller;

      const bundle = await controller.waitForBundle((_, elapsed) => {
        setStatusText(`Running "${sheetTitle}"... (${elapsed}s)`);
      });
      const dataRows = bundle.rows;

      const expandedColCount = expandPivotedColumns(
        columns,
        dataRows,
        orderedPayload.pivots || undefined,
        Boolean(orderedPayload.row_total),
        bundle.pivotsMeta,
        tableOptions.pivotGroupMode || "measure_first"
      ).expandedColumns.length;

      const effectiveDest = forceNewSheet ? "new" : destination;
      const effectiveAnchor =
        effectiveDest === "new" ? "A1" : (tableOptions.anchorCell || "A1").toUpperCase();
      const effectiveTableId =
        effectiveDest === "new"
          ? "Looker_A1"
          : tableOptions.tableId || `Looker_${effectiveAnchor}`;
      const siblingRanges =
        effectiveDest === "active"
          ? getSiblingTableBoundingBoxes(allSheetTables, effectiveTableId, effectiveAnchor)
          : [];

      const writeOptions: WriteOptions = {
        ...tableOptions,
        anchorCell: effectiveAnchor,
        tableId: effectiveTableId,
        siblingTableRanges: siblingRanges,
        destination: effectiveDest,
        sheetName: sheetTitle,
        pivots: orderedPayload.pivots || undefined,
        pivotsMeta: bundle.pivotsMeta,
        includeTotals: Boolean(orderedPayload.total),
        includeRowTotals: Boolean(orderedPayload.row_total),
        includeSubtotals: Boolean(orderedPayload.subtotals && orderedPayload.subtotals.length > 0),
        totalsData: bundle.totalsData,
        subtotalsData: bundle.subtotalsData,
      };

      const writtenSheetId = await writeLookerDataToWorksheet(
        columns,
        dataRows,
        writeOptions,
        (pct) => setProgressPercent(pct)
      );

      const totalWrittenDataRows =
        dataRows.length +
        (bundle.subtotalsData ? Object.values(bundle.subtotalsData).flat().length : 0) +
        (bundle.totalsData ? 1 : 0);

      const config: StoredSheetConfig = {
        tableId: effectiveTableId,
        anchorCell: effectiveAnchor,
        queryPayload: orderedPayload,
        columns,
        options: {
          ...writeOptions,
          totalsData: undefined,
          subtotalsData: undefined,
          siblingTableRanges: undefined,
          destination: "active",
          targetSheetId: undefined,
          prevRowCount: totalWrittenDataRows,
          prevColCount: expandedColCount,
        },
        lastRefreshed: new Date().toISOString(),
        rowCount: totalWrittenDataRows,
        modelLabel: payload.model,
        exploreLabel: sheetTitle || expMeta?.label || payload.view,
      };

      await saveSheetMetadata(config, writtenSheetId);
      preferredTableIdRef.current = effectiveTableId;
      if (writtenSheetId) {
        sourceSheetIdRef.current = writtenSheetId;
        setSourceSheetId(writtenSheetId);
      }
      if (!forceNewSheet) {
        await checkSheetMetadata(true);
      }
    } catch (e: any) {
      setErrorMessage(e.message || "Failed to import Look / Dashboard tile.");
    } finally {
      setIsExecuting(false);
      activeTaskRef.current = null;
    }
  };

  return (
    <div className={styles.appRoot}>
      <Header
        isConnected={!!auth}
        instanceUrl={auth?.baseUrl}
        userName={user?.first_name ? `${user.first_name} ${user.last_name || ""}` : undefined}
        onDisconnect={handleDisconnect}
        onRefreshAll={handleRefreshAll}
        refreshingAll={refreshingAll}
        activeMode={navMode}
        onModeChange={(mode) => {
          setNavMode(mode);
          if (mode !== "explore" && activeView === "refresh") {
            setActiveView("builder");
          }
        }}
      />

      {errorMessage && (
        <MessageBar intent="error" style={{ margin: "8px 16px" }}>
          <MessageBarBody>{errorMessage}</MessageBarBody>
          <MessageBarActions
            containerAction={
              <Button
                appearance="subtle"
                aria-label="Dismiss error"
                icon={<DismissRegular />}
                onClick={() => setErrorMessage(null)}
              />
            }
          />
        </MessageBar>
      )}

      {!auth ? (
        <AuthView onSuccess={handleAuthSuccess} />
      ) : navMode === "looks" ? (
        <div className={styles.scrollContent}>
          <LooksBrowser
            baseUrl={auth.baseUrl}
            token={auth.accessToken}
            isExecuting={isExecuting}
            onImportPayload={(p, title) => handleImportExternalPayload(p, title)}
            onOpenInExplore={handleOpenPayloadInExplore}
            onError={setErrorMessage}
          />
        </div>
      ) : navMode === "dashboards" ? (
        <div className={styles.scrollContent}>
          <DashboardsBrowser
            baseUrl={auth.baseUrl}
            token={auth.accessToken}
            isExecuting={isExecuting}
            onImportPayload={(p, title, forceNew) =>
              handleImportExternalPayload(p, title, forceNew)
            }
            onOpenInExplore={handleOpenPayloadInExplore}
            onError={setErrorMessage}
          />
        </div>
      ) : activeView === "refresh" && sheetConfig ? (
        <RefreshView
          config={sheetConfig}
          allSheetTables={allSheetTables}
          exploreDimensions={exploreDetail?.fields.dimensions || []}
          onSelectTable={(tbl) => {
            preferredTableIdRef.current = tbl.tableId || null;
            setSheetConfig(tbl);
          }}
          onAddTableToSheet={handleAddTableToSheet}
          onApplyDrillPayload={(updatedPayload, updatedOpts) =>
            handleRefreshActive(updatedPayload, updatedOpts)
          }
          onChangeDisplayLevel={async (nextLevel: RowGroupDisplayLevel) => {
            const updatedNodes = await applyWorksheetTreeDisplayLevel(
              sourceSheetId || undefined,
              sheetConfig.options,
              sheetConfig.rowCount,
              nextLevel
            );
            if (updatedNodes) {
              const updatedCfg: StoredSheetConfig = {
                ...sheetConfig,
                options: {
                  ...sheetConfig.options,
                  defaultDisplayLevel: nextLevel,
                  treeNodes: updatedNodes,
                },
              };
              const nextTables = await saveSheetMetadata(
                updatedCfg,
                sourceSheetId || undefined
              );
              setAllSheetTables(nextTables);
              setSheetConfig(updatedCfg);
            }
          }}
          onRefresh={() => handleRefreshActive()}
          onCancel={handleCancelQuery}
          onEdit={() => handleEditQuery(sheetConfig)}
          onNewQuery={handleStartNewQuery}
          isRefreshing={isExecuting}
          statusText={statusText}
          progressPercent={progressPercent}
        />
      ) : (
        <>
          <div className={styles.scrollContent}>
            <ModelSelector
              models={models}
              selectedModel={selectedModel}
              selectedExplore={selectedExplore}
              onModelChange={(m) => {
                setSelectedModel(m);
                setSelectedFields([]);
                setSorts([]);
                setPivots([]);
                setFilters([]);
                setPromptValues({});
                const modelObj = models.find((mod) => mod.name === m);
                if (modelObj && modelObj.explores.length > 0) {
                  setSelectedExplore(modelObj.explores[0].name);
                } else {
                  setSelectedExplore("");
                  setExploreDetail(null);
                }
              }}
              onExploreChange={(exp) => {
                setSelectedExplore(exp);
                setSelectedFields([]);
                setSorts([]);
                setPivots([]);
                setFilters([]);
                setPromptValues({});
              }}
              loadingModels={loadingModels}
              loadingExplore={loadingExplore}
            />

            {exploreDetail && (
              <>
                <PromptsDialog
                  parameters={exploreDetail.fields.parameters}
                  alwaysFilters={exploreDetail.always_filter}
                  fields={[...exploreDetail.fields.dimensions, ...exploreDetail.fields.measures]}
                  promptValues={promptValues}
                  onPromptChange={(k, v) => setPromptValues((prev) => ({ ...prev, [k]: v }))}
                  baseUrl={auth?.baseUrl}
                  token={auth?.accessToken}
                  modelName={selectedModel}
                  exploreName={selectedExplore}
                />

                <FieldPicker
                  dimensions={exploreDetail.fields.dimensions}
                  measures={exploreDetail.fields.measures}
                  selectedFields={selectedFields}
                  onToggleField={handleToggleField}
                  onClearSelected={handleClearSelected}
                  sorts={sorts}
                  onToggleSort={handleToggleSort}
                  pivots={pivots}
                  onTogglePivot={handleTogglePivot}
                  onSetPivots={setPivots}
                  onReorderFields={setSelectedFields}
                  options={tableOptions}
                  onOptionsChange={(upd) => setTableOptions((prev) => ({ ...prev, ...upd }))}
                />

                <FilterBuilder
                  fields={[...exploreDetail.fields.dimensions, ...exploreDetail.fields.measures]}
                  filters={filters}
                  onAddFilter={handleAddFilter}
                  onUpdateFilter={handleUpdateFilter}
                  onRemoveFilter={handleRemoveFilter}
                  baseUrl={auth?.baseUrl}
                  token={auth?.accessToken}
                  modelName={selectedModel}
                  exploreName={selectedExplore}
                />

                <TableSettings
                  rowLimit={rowLimit}
                  onRowLimitChange={setRowLimit}
                  destination={destination}
                  onDestinationChange={setDestination}
                  options={tableOptions}
                  onOptionsChange={(upd) => setTableOptions((prev) => ({ ...prev, ...upd }))}
                />
              </>
            )}
          </div>

          <ExecutionBar
            isExecuting={isExecuting}
            statusText={statusText}
            progressPercent={progressPercent}
            canExecute={selectedFields.length > 0 && !loadingExplore && !isExecuting}
            onExecute={handleExecuteImport}
            onCancel={handleCancelQuery}
          />
        </>
      )}
    </div>
  );
};
