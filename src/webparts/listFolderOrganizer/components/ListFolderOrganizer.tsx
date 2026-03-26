import * as React from "react";
import {
  DetailsList,
  IColumn,
  SelectionMode,
  DetailsListLayoutMode,
  Checkbox,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Icon,
  ProgressIndicator,
  Stack,
  Text,
} from "@fluentui/react";
import { IListFolderOrganizerProps } from "./IListFolderOrganizerProps";
import {
  IListInfo,
  ListProcessingStatus,
} from "../models/IListInfo";
import {
  GroupingStrategy,
  MaxLevelsForStrategy,
} from "../models/GroupingType";
import { SharePointService } from "../services/SharePointService";
import styles from "./ListFolderOrganizer.module.scss";

/* ------------------------------------------------------------------ */
/*  Dropdown option helpers                                            */
/* ------------------------------------------------------------------ */

const strategyOptions: IDropdownOption[] = [
  { key: GroupingStrategy.Date, text: "By Date" },
  { key: GroupingStrategy.Text, text: "By Initial Letter" },
  { key: GroupingStrategy.Choice, text: "By Choice Value" },
];

function buildLevelsOptions(strategy: GroupingStrategy): IDropdownOption[] {
  const max = MaxLevelsForStrategy[strategy] || 1;
  const opts: IDropdownOption[] = [];
  for (let i = 1; i <= max; i++) {
    let label = i + " level" + (i > 1 ? "s" : "");
    if (strategy === GroupingStrategy.Date) {
      if (i === 1) label += " (Year)";
      if (i === 2) label += " (Year / Month)";
      if (i === 3) label += " (Year / Month / Day)";
    } else if (strategy === GroupingStrategy.Text) {
      if (i === 1) label += " (1st letter)";
      if (i === 2) label += " (1st + 2nd letter)";
      if (i === 3) label += " (1st + 2nd + 3rd letter)";
    }
    opts.push({ key: i, text: label });
  }
  return opts;
}

/* ================================================================== */
/*  Component                                                          */
/* ================================================================== */

const ListFolderOrganizer: React.FC<IListFolderOrganizerProps> = (props) => {
  const [lists, setLists] = React.useState<IListInfo[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>("");
  const [processing, setProcessing] = React.useState<boolean>(false);
  const [logMessages, setLogMessages] = React.useState<string[]>([]);

  const service = React.useMemo(
    () => new SharePointService(props.spHttpClient, props.siteUrl),
    [props.spHttpClient, props.siteUrl]
  );

  /* ----- initial load ----- */
  React.useEffect(() => {
    loadLists();
  }, []);

  const loadLists = async (): Promise<void> => {
    setLoading(true);
    setError("");
    try {
      const result = await service.getLists();
      setLists(result);

      // Fetch root item counts in the background
      result.forEach((list) => {
        service
          .getRootItemCount(list.id, list.rootFolderUrl)
          .then((count) => {
            setLists((prev) =>
              prev.map((l) =>
                l.id === list.id ? { ...l, rootItemCount: count } : l
              )
            );
          })
          .catch(() => {
            /* swallow – count stays at 0 */
          });
      });
    } catch (err: any) {
      setError("Failed to load lists: " + (err.message || err));
    } finally {
      setLoading(false);
    }
  };

  /* ----- row-level handlers ----- */

  /** Toggle the checkbox and lazy-load fields for the current strategy. */
  const onSelectToggle = async (
    listId: string,
    checked: boolean
  ): Promise<void> => {
    setLists((prev) =>
      prev.map((l) => (l.id === listId ? { ...l, selected: checked } : l))
    );

    if (checked) {
      const list = lists.filter((l) => l.id === listId)[0];
      if (list && list.availableFields.length === 0) {
        try {
          const fields = await service.getListFields(
            listId,
            list.groupingStrategy
          );
          setLists((prev) =>
            prev.map((l) =>
              l.id === listId ? { ...l, availableFields: fields } : l
            )
          );
        } catch (err: any) {
          setError("Failed to load fields: " + (err.message || err));
        }
      }
    }
  };

  /** When the strategy changes, reload matching fields. */
  const onStrategyChange = async (
    listId: string,
    strategy: GroupingStrategy
  ): Promise<void> => {
    setLists((prev) =>
      prev.map((l) =>
        l.id === listId
          ? {
              ...l,
              groupingStrategy: strategy,
              sourceFieldInternalName: "",
              sourceFieldTitle: "",
              availableFields: [],
              levels: 1,
            }
          : l
      )
    );

    try {
      const fields = await service.getListFields(listId, strategy);
      setLists((prev) =>
        prev.map((l) =>
          l.id === listId ? { ...l, availableFields: fields } : l
        )
      );
    } catch (err: any) {
      setError("Failed to load fields: " + (err.message || err));
    }
  };

  const onLevelsChange = (listId: string, levels: number): void => {
    setLists((prev) =>
      prev.map((l) => (l.id === listId ? { ...l, levels: levels } : l))
    );
  };

  const onSourceFieldChange = (
    listId: string,
    fieldInternalName: string,
    fieldTitle: string
  ): void => {
    setLists((prev) =>
      prev.map((l) =>
        l.id === listId
          ? {
              ...l,
              sourceFieldInternalName: fieldInternalName,
              sourceFieldTitle: fieldTitle,
            }
          : l
      )
    );
  };

  /* ----- logging helpers ----- */
  const addLog = (msg: string): void => {
    setLogMessages((prev) => [
      ...prev,
      "[" + new Date().toLocaleTimeString() + "] " + msg,
    ]);
  };

  const updateListStatus = (
    listId: string,
    status: ListProcessingStatus,
    message: string,
    progress: number
  ): void => {
    setLists((prev) =>
      prev.map((l) =>
        l.id === listId
          ? { ...l, status: status, statusMessage: message, progress: progress }
          : l
      )
    );
  };

  /* ----- main action ----- */

  const organizeSelected = async (): Promise<void> => {
    const selected = lists.filter(
      (l) => l.selected && l.sourceFieldInternalName
    );
    if (selected.length === 0) {
      setError(
        "Select at least one list and configure its grouping (strategy + source field)."
      );
      return;
    }

    setProcessing(true);
    setLogMessages([]);
    setError("");

    for (const list of selected) {
      try {
        updateListStatus(
          list.id,
          ListProcessingStatus.Processing,
          "Starting…",
          0
        );
        addLog("Processing list: " + list.title);

        /* 1. Enable folders if needed */
        if (!list.folderCreationEnabled) {
          addLog('  Enabling folder creation for "' + list.title + '"…');
          await service.enableFolderCreation(list.id);
          setLists((prev) =>
            prev.map((l) =>
              l.id === list.id ? { ...l, folderCreationEnabled: true } : l
            )
          );
          addLog("  ✓ Folder creation enabled");
        }

        updateListStatus(
          list.id,
          ListProcessingStatus.Processing,
          "Fetching items…",
          5
        );

        /* 2. Get all items */
        addLog("  Fetching items…");
        const items = await service.getListItems(
          list.id,
          list.sourceFieldInternalName
        );
        addLog("  Found " + items.length + " items");

        if (items.length === 0) {
          updateListStatus(
            list.id,
            ListProcessingStatus.Done,
            "No items to organize.",
            100
          );
          addLog("  No items found — skipping.");
          continue;
        }

        /* 3. Calculate folder paths */
        const folderPaths: string[] = [];
        const seen: { [key: string]: boolean } = {};
        const itemMoves: Array<{ itemId: number; folderPath: string }> = [];

        for (var idx = 0; idx < items.length; idx++) {
          var item = items[idx];
          var folderPath = SharePointService.generateFolderPath(
            item,
            list.sourceFieldInternalName,
            list.groupingStrategy,
            list.levels
          );
          if (!seen[folderPath]) {
            seen[folderPath] = true;
            folderPaths.push(folderPath);
          }
          var targetDir = list.rootFolderUrl + "/" + folderPath;
          if (item.FileDirRef !== targetDir) {
            itemMoves.push({ itemId: item.Id, folderPath: folderPath });
          }
        }

        /* 4. Create folders */
        addLog("  Creating " + folderPaths.length + " folder(s)…");
        for (var fi = 0; fi < folderPaths.length; fi++) {
          await service.ensureFolder(list.rootFolderUrl, folderPaths[fi]);
          updateListStatus(
            list.id,
            ListProcessingStatus.Processing,
            "Creating folders (" + (fi + 1) + "/" + folderPaths.length + ")…",
            5 + Math.round(((fi + 1) / folderPaths.length) * 25)
          );
        }
        addLog("  ✓ " + folderPaths.length + " folder(s) ensured");

        /* 5. Move items */
        if (itemMoves.length === 0) {
          addLog("  All items already in correct folders.");
          updateListStatus(
            list.id,
            ListProcessingStatus.Done,
            "All items already organized.",
            100
          );
          continue;
        }

        addLog("  Moving " + itemMoves.length + " item(s) to folders…");
        var moveCount = 0;
        for (var mi = 0; mi < itemMoves.length; mi++) {
          var move = itemMoves[mi];
          var dest = list.rootFolderUrl + "/" + move.folderPath;
          try {
            await service.moveItemToFolder(list.id, move.itemId, dest);
          } catch (moveErr: any) {
            addLog(
              "  ⚠ Failed to move item " +
                move.itemId +
                ": " +
                (moveErr.message || moveErr)
            );
          }
          moveCount++;
          if (moveCount % 10 === 0 || moveCount === itemMoves.length) {
            updateListStatus(
              list.id,
              ListProcessingStatus.Processing,
              "Moving items (" +
                moveCount +
                "/" +
                itemMoves.length +
                ")…",
              30 + Math.round((moveCount / itemMoves.length) * 70)
            );
          }
        }

        addLog("  ✓ " + moveCount + " item(s) organized");
        updateListStatus(
          list.id,
          ListProcessingStatus.Done,
          "Done! " + moveCount + " items organized.",
          100
        );
      } catch (err: any) {
        addLog("  ✗ Error: " + (err.message || err));
        updateListStatus(
          list.id,
          ListProcessingStatus.Error,
          err.message || String(err),
          0
        );
      }
    }

    setProcessing(false);
    addLog("— All done! —");
  };

  /* ----- Select / Deselect All ----- */
  const selectAll = (): void => {
    setLists((prev) => prev.map((l) => ({ ...l, selected: true })));
    // lazy-load fields for every list that doesn't have them yet
    lists.forEach((list) => {
      if (list.availableFields.length === 0) {
        service
          .getListFields(list.id, list.groupingStrategy)
          .then((fields) => {
            setLists((prev) =>
              prev.map((l) =>
                l.id === list.id ? { ...l, availableFields: fields } : l
              )
            );
          })
          .catch(() => {
            /* swallow individual failures */
          });
      }
    });
  };

  const deselectAll = (): void => {
    setLists((prev) => prev.map((l) => ({ ...l, selected: false })));
  };

  /* ================================================================ */
  /*  Column definitions                                               */
  /* ================================================================ */

  const columns: IColumn[] = [
    {
      key: "selected",
      name: "",
      minWidth: 32,
      maxWidth: 32,
      onRender: (item: IListInfo) => (
        <Checkbox
          checked={item.selected}
          onChange={(_ev, checked) => onSelectToggle(item.id, !!checked)}
          disabled={processing}
        />
      ),
    },
    {
      key: "title",
      name: "List Name",
      fieldName: "title",
      minWidth: 140,
      maxWidth: 250,
      isResizable: true,
    },
    {
      key: "itemCount",
      name: "Total",
      fieldName: "itemCount",
      minWidth: 50,
      maxWidth: 60,
    },
    {
      key: "rootItemCount",
      name: "In Root",
      minWidth: 60,
      maxWidth: 75,
      onRender: (item: IListInfo) => {
        if (item.rootItemCount === -1) {
          return <Text title="List exceeds 5 000-item view threshold">&gt; 5 000</Text>;
        }
        return <Text>{item.rootItemCount}</Text>;
      },
    },
    {
      key: "folderEnabled",
      name: "Folders",
      minWidth: 55,
      maxWidth: 55,
      onRender: (item: IListInfo) => (
        <Icon
          iconName={item.folderCreationEnabled ? "CheckMark" : "Cancel"}
          className={
            item.folderCreationEnabled
              ? styles.folderEnabled
              : styles.folderDisabled
          }
          title={
            item.folderCreationEnabled
              ? "Folder creation enabled"
              : "Folder creation disabled – will be auto-enabled"
          }
        />
      ),
    },
    {
      key: "strategy",
      name: "Grouping",
      minWidth: 140,
      maxWidth: 170,
      onRender: (item: IListInfo) => (
        <Dropdown
          selectedKey={item.groupingStrategy}
          options={strategyOptions}
          onChange={(_ev, option) =>
            option &&
            onStrategyChange(item.id, option.key as GroupingStrategy)
          }
          disabled={processing}
          className={styles.cellDropdown}
        />
      ),
    },
    {
      key: "levels",
      name: "Depth",
      minWidth: 165,
      maxWidth: 210,
      onRender: (item: IListInfo) => (
        <Dropdown
          selectedKey={item.levels}
          options={buildLevelsOptions(item.groupingStrategy)}
          onChange={(_ev, option) =>
            option && onLevelsChange(item.id, option.key as number)
          }
          disabled={processing}
          className={styles.cellDropdown}
        />
      ),
    },
    {
      key: "sourceField",
      name: "Source Field",
      minWidth: 160,
      maxWidth: 220,
      onRender: (item: IListInfo) => (
        <Dropdown
          selectedKey={item.sourceFieldInternalName || undefined}
          options={item.availableFields.map((f) => ({
            key: f.internalName,
            text: f.title + " (" + f.typeAsString + ")",
          }))}
          placeholder={
            item.availableFields.length === 0
              ? "Check row to load fields"
              : "Select field…"
          }
          onChange={(_ev, option) => {
            if (option) {
              var field = item.availableFields.filter(
                (f) => f.internalName === option.key
              )[0];
              onSourceFieldChange(
                item.id,
                option.key as string,
                field ? field.title : ""
              );
            }
          }}
          disabled={processing || item.availableFields.length === 0}
          className={styles.cellDropdown}
        />
      ),
    },
    {
      key: "status",
      name: "Status",
      minWidth: 180,
      maxWidth: 280,
      onRender: (item: IListInfo) => {
        if (item.status === ListProcessingStatus.Processing) {
          return (
            <Stack>
              <ProgressIndicator
                percentComplete={item.progress / 100}
                description={item.statusMessage}
              />
            </Stack>
          );
        }
        if (item.status === ListProcessingStatus.Done) {
          return (
            <Text className={styles.statusDone}>
              <Icon iconName="CheckMark" /> {item.statusMessage}
            </Text>
          );
        }
        if (item.status === ListProcessingStatus.Error) {
          return (
            <Text className={styles.statusError}>
              <Icon iconName="ErrorBadge" /> {item.statusMessage}
            </Text>
          );
        }
        return <Text className={styles.statusIdle}>Ready</Text>;
      },
    },
  ];

  /* ================================================================ */
  /*  Render                                                           */
  /* ================================================================ */

  return (
    <div className={styles.listFolderOrganizer}>
      <Stack tokens={{ childrenGap: 16 }}>
        {/* Header */}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          wrap
          tokens={{ childrenGap: 8 }}
        >
          <Text variant="xLarge" className={styles.title}>
            <Icon iconName="FolderList" /> List Folder Organizer
          </Text>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Select All"
              onClick={selectAll}
              disabled={processing || loading}
            />
            <DefaultButton
              text="Deselect All"
              onClick={deselectAll}
              disabled={processing || loading}
            />
            <DefaultButton
              iconProps={{ iconName: "Refresh" }}
              text="Refresh"
              onClick={loadLists}
              disabled={processing}
            />
            <PrimaryButton
              iconProps={{ iconName: "Play" }}
              text="Organize Selected"
              onClick={organizeSelected}
              disabled={processing || loading}
            />
          </Stack>
        </Stack>

        {/* Error banner */}
        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setError("")}
            isMultiline={false}
          >
            {error}
          </MessageBar>
        )}

        {/* Main content */}
        {loading ? (
          <Spinner size={SpinnerSize.large} label="Loading lists…" />
        ) : lists.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No custom lists found on this site.
          </MessageBar>
        ) : (
          <div className={styles.tableContainer}>
            <DetailsList
              items={lists}
              columns={columns}
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              compact={false}
            />
          </div>
        )}

        {/* Activity log */}
        {logMessages.length > 0 && (
          <div className={styles.logPanel}>
            <Text
              variant="smallPlus"
              style={{ fontWeight: 600, marginBottom: 4, display: "block" }}
            >
              Activity Log
            </Text>
            {logMessages.map((msg, i) => (
              <div key={i} className={styles.logLine}>
                {msg}
              </div>
            ))}
          </div>
        )}
      </Stack>
    </div>
  );
};

export default ListFolderOrganizer;
