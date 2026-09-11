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
import { Header } from "./Header";
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

import { getStoredAuth, clearAuth, ensureValidToken, AuthTokens } from "../services/lookerAuth";
import {
  getModels,
  getExplore,
  getCurrentUser,
  ModelSummary,
  ExploreDetail,
  UserSummary,
} from "../services/lookerClient";
import { resolveExcelNumberFormat } from "../services/formatMapper";
import { runQueryTask, QueryTaskController, QueryPayload } from "../services/queryTaskRunner";
import {
  writeLookerDataToWorksheet,
  ColumnDefinition,
  WriteOptions,
} from "../services/excelWriter";
import {
  getActiveSheetMetadata,
  saveSheetMetadata,
  getAllWorkbookLookerSheets,
  StoredSheetConfig,
} from "../services/sheetMetadata";

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

  // View state: 'builder' | 'refresh'
  const [activeView, setActiveView] = useState<"builder" | "refresh">("builder");
  const [sheetConfig, setSheetConfig] = useState<StoredSheetConfig | null>(null);

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
  const [rowLimit, setRowLimit] = useState<string>("5000");
  const [destination, setDestination] = useState<"active" | "new">("active");
  const [tableOptions, setTableOptions] = useState<WriteOptions>({
    useExcelTable: true,
    tableStyle: "TableStyleMedium2",
    preserveUserFormatting: true,
    autoExpandFormulas: true,
    freezeHeader: true,
    nullDisplay: "",
  });

  // Execution & Progress state
  const [isExecuting, setIsExecuting] = useState(false);
  const [statusText, setStatusText] = useState("");
  const [progressPercent, setProgressPercent] = useState(0);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [refreshingAll, setRefreshingAll] = useState(false);

  const activeTaskRef = useRef<QueryTaskController | null>(null);

  // Check active sheet metadata
  const checkSheetMetadata = useCallback(async () => {
    const meta = await getActiveSheetMetadata();
    setSheetConfig(meta);
    if (meta) {
      setActiveView("refresh");
    } else {
      setActiveView("builder");
    }
  }, []);

  // Listen for worksheet switches
  useEffect(() => {
    checkSheetMetadata();

    const onSelectionChange = () => {
      checkSheetMetadata();
    };

    if (window.Office && Office.context && Office.context.document) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChange
      );
    }

    return () => {
      if (window.Office && Office.context && Office.context.document) {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          onSelectionChange
        );
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
        // Default to first model and explore if available
        const defaultModel = data[0];
        if (defaultModel) {
          setSelectedModel(defaultModel.name);
          const defaultExp = defaultModel.explores?.[0];
          if (defaultExp) {
            setSelectedExplore(defaultExp.name);
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

        // Pre-fill defaults for any required always_filter
        if (data.always_filter && data.always_filter.length > 0) {
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
    setSelectedModel(config.queryPayload.model);
    setSelectedExplore(config.queryPayload.view);
    setSelectedFields(config.queryPayload.fields || []);
    setSorts(config.queryPayload.sorts || []);
    setPivots(config.queryPayload.pivots || []);
    setRowLimit(config.queryPayload.limit || "5000");
    if (config.options) {
      setTableOptions(config.options);
      setDestination("active");
    }

    if (config.queryPayload.filters) {
      const restoredPrompts: Record<string, string> = {};
      const restoredFilters: FilterCondition[] = [];
      Object.entries(config.queryPayload.filters).forEach(([field, expr], idx) => {
        restoredPrompts[field] = expr;
        restoredFilters.push({
          id: `${Date.now()}_${idx}`,
          field,
          operator: "is",
          value: expr,
        });
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
    setSelectedFields([]);
    setSorts([]);
    setPivots([]);
    setFilters([]);
    setPromptValues({});
    setDestination("new");
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
            compiledFilters[f.field] = expr;
          }
        }
      });

      const queryPayload: QueryPayload = {
        model: selectedModel,
        view: selectedExplore,
        fields: selectedFields,
        filters: Object.keys(compiledFilters).length > 0 ? compiledFilters : null,
        sorts: sorts.length > 0 ? sorts : [],
        limit: rowLimit,
        pivots: pivots.length > 0 ? pivots : null,
      };

      // 2. Prepare column metadata with number formats
      const columns: ColumnDefinition[] = selectedFields.map((fieldKey) => {
        const def = allAvailableFields.find((f) => f.name === fieldKey);
        return {
          fieldKey,
          label: def?.label || fieldKey,
          type: def?.type,
          excelFormat: def ? resolveExcelNumberFormat(def) : undefined,
        };
      });

      // 3. Launch Async Query Task
      const controller = await runQueryTask(auth.baseUrl, auth.accessToken, queryPayload);
      activeTaskRef.current = controller;

      const dataRows = await controller.waitForResults((status, elapsed) => {
        setStatusText(`Running query in warehouse... (${elapsed}s)`);
      });

      // 4. Stream to Excel
      setStatusText(`Streaming ${dataRows.length.toLocaleString()} rows into worksheet...`);
      const writeOptions: WriteOptions = {
        ...tableOptions,
        destination,
        sheetName: exploreDetail.label || selectedExplore || "Looker Data",
      };
      await writeLookerDataToWorksheet(columns, dataRows, writeOptions, (pct) => {
        setProgressPercent(pct);
      });

      // 5. Persist metadata for one-click refresh
      const config: StoredSheetConfig = {
        queryPayload,
        columns,
        options: writeOptions,
        lastRefreshed: new Date().toISOString(),
        rowCount: dataRows.length,
        modelLabel: models.find((m) => m.name === selectedModel)?.label,
        exploreLabel: exploreDetail.label,
      };
      await saveSheetMetadata(config);
      setSheetConfig(config);
      setActiveView("refresh");
    } catch (err: any) {
      setErrorMessage(err.message || "Query execution failed.");
    } finally {
      setIsExecuting(false);
      activeTaskRef.current = null;
    }
  };

  // Re-run Refresh on active sheet
  const handleRefreshActive = async () => {
    if (!auth || !sheetConfig) return;

    setIsExecuting(true);
    setStatusText("Submitting refresh query...");
    setProgressPercent(0);
    setErrorMessage(null);

    try {
      const validAuth = await ensureValidToken(auth);
      if (validAuth.accessToken !== auth.accessToken) {
        setAuth(validAuth);
      }

      const controller = await runQueryTask(
        validAuth.baseUrl,
        validAuth.accessToken,
        sheetConfig.queryPayload
      );
      activeTaskRef.current = controller;

      const dataRows = await controller.waitForResults((_, elapsed) => {
        setStatusText(`Refreshing data from warehouse... (${elapsed}s)`);
      });

      setStatusText(`Writing ${dataRows.length.toLocaleString()} rows...`);
      await writeLookerDataToWorksheet(
        sheetConfig.columns,
        dataRows,
        { ...sheetConfig.options, destination: "active" },
        (pct) => setProgressPercent(pct)
      );

      const updatedConfig: StoredSheetConfig = {
        ...sheetConfig,
        lastRefreshed: new Date().toISOString(),
        rowCount: dataRows.length,
      };
      await saveSheetMetadata(updatedConfig);
      setSheetConfig(updatedConfig);
    } catch (err: any) {
      setErrorMessage(err.message || "Failed to refresh data.");
    } finally {
      setIsExecuting(false);
      activeTaskRef.current = null;
    }
  };

  // Global "Refresh All" across workbook
  const handleRefreshAll = async () => {
    if (!auth) return;
    setRefreshingAll(true);
    try {
      const sheets = await getAllWorkbookLookerSheets();
      for (const s of sheets) {
        const controller = await runQueryTask(auth.baseUrl, auth.accessToken, s.config.queryPayload);
        const dataRows = await controller.waitForResults(() => {});
        await writeLookerDataToWorksheet(s.config.columns, dataRows, s.config.options, () => {});
        await saveSheetMetadata({
          ...s.config,
          lastRefreshed: new Date().toISOString(),
          rowCount: dataRows.length,
        });
      }
      await checkSheetMetadata();
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

  return (
    <div className={styles.appRoot}>
      <Header
        isConnected={!!auth}
        instanceUrl={auth?.baseUrl}
        userName={user?.first_name ? `${user.first_name} ${user.last_name || ""}` : undefined}
        onDisconnect={handleDisconnect}
        onRefreshAll={handleRefreshAll}
        refreshingAll={refreshingAll}
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
      ) : activeView === "refresh" && sheetConfig ? (
        <RefreshView
          config={sheetConfig}
          onRefresh={handleRefreshActive}
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
                  promptValues={promptValues}
                  onPromptChange={(k, v) => setPromptValues((prev) => ({ ...prev, [k]: v }))}
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
                  onReorderFields={setSelectedFields}
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
