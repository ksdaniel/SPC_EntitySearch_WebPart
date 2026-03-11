import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  SearchBox,
  Button,
  Badge,
  Text,
} from '@fluentui/react-components';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './EntitySearchWebpart.module.scss';
import type { IEntitySearchWebpartProps } from './IEntitySearchWebpartProps';

interface IEntity {
  id: number;
  name: string;
  type: string;
  deal: string;
  status: string;
  fields: Record<string, string>;
}

const DEFAULT_PAGE_SIZE = 5;

interface IActionLink {
  label: string;
  url: string;
}

interface ISingleRowAction {
  type: 'single';
  label: string;
  url: string;
}

interface IDropdownRowAction {
  type: 'dropdown';
  label: string;
  items: IActionLink[];
}

type RowAction = ISingleRowAction | IDropdownRowAction;

const EntitySearchWebpart: React.FC<IEntitySearchWebpartProps> = (props) => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [openMenuKey, setOpenMenuKey] = React.useState<string | null>(null);
  const [currentPage, setCurrentPage] = React.useState<number>(1);
  const [entities, setEntities] = React.useState<IEntity[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');

  const configuredActionFields = React.useMemo(
    () => extractTemplateFields(props.actionsConfigurationJson),
    [props.actionsConfigurationJson]
  );

  const { actions, parseError } = React.useMemo(
    () => parseActionsConfiguration(props.actionsConfigurationJson),
    [props.actionsConfigurationJson]
  );

  React.useEffect(() => {
    let isActive = true;

    const loadEntities = async (): Promise<void> => {
      if (!props.listId) {
        setEntities([]);
        setError('Select a list in the web part settings to start searching entities.');
        return;
      }

      if (!props.titleFieldInternalName) {
        setEntities([]);
        setError('Map the title field in the web part settings.');
        return;
      }

      setIsLoading(true);
      setError('');

      try {
        const fieldsToSelect = [
          'Id',
          props.titleFieldInternalName,
          props.typeFieldInternalName,
          props.dealFieldInternalName,
          props.statusFieldInternalName,
          ...configuredActionFields
        ]
          .filter((value, index, self) => !!value && self.indexOf(value) === index)
          .join(',');

        const endpoint = `${props.siteUrl}/_api/web/lists(guid'${props.listId}')/items?$select=${fieldsToSelect}&$top=200`;
        const response = await props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

        if (!response.ok) {
          throw new Error(`Failed to load entities: ${response.statusText}`);
        }

        const data = await response.json() as { value: Array<Record<string, unknown>> };
        const mappedEntities = data.value.map((item) => ({
          id: Number(item.Id),
          name: readFieldValue(item, props.titleFieldInternalName),
          type: readFieldValue(item, props.typeFieldInternalName),
          deal: readFieldValue(item, props.dealFieldInternalName),
          status: readFieldValue(item, props.statusFieldInternalName),
          fields: mapItemFields(item)
        }));

        if (isActive) {
          setEntities(mappedEntities);
        }
      } catch (loadError) {
        if (isActive) {
          setEntities([]);
          setError('Unable to load entities. Verify list and field mappings in web part settings.');
        }
        console.error('EntitySearchWebPart: failed to load entities.', loadError);
      } finally {
        if (isActive) {
          setIsLoading(false);
        }
      }
    };

    loadEntities().catch((loadError) => {
      console.error('EntitySearchWebPart: failed to initialize entity load.', loadError);
    });

    return () => {
      isActive = false;
    };
  }, [
    props.listId,
    props.titleFieldInternalName,
    props.typeFieldInternalName,
    props.dealFieldInternalName,
    props.statusFieldInternalName,
    configuredActionFields,
    props.siteUrl,
    props.spHttpClient
  ]);

  const normalizedQuery = searchQuery.trim().toLowerCase();

  const filtered = entities.filter(e =>
    e.name.toLowerCase().includes(normalizedQuery) ||
    e.type.toLowerCase().includes(normalizedQuery) ||
    e.deal.toLowerCase().includes(normalizedQuery)
  );

  const totalPages = Math.max(1, Math.ceil(filtered.length / DEFAULT_PAGE_SIZE));

  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [currentPage, totalPages]);

  const pageStartIndex = (currentPage - 1) * DEFAULT_PAGE_SIZE;
  const pagedEntities = filtered.slice(pageStartIndex, pageStartIndex + DEFAULT_PAGE_SIZE);
  const pageStartNumber = filtered.length > 0 ? pageStartIndex + 1 : 0;
  const pageEndNumber = pageStartIndex + pagedEntities.length;

  // Close menu when clicking outside
  React.useEffect(() => {
    const handleClickOutside = (): void => setOpenMenuKey(null);
    document.addEventListener('click', handleClickOutside);
    return () => document.removeEventListener('click', handleClickOutside);
  }, []);

  const onActionClick = React.useCallback((actionUrl: string, entity: IEntity): void => {
    const resolvedUrl = resolveActionUrl(actionUrl, entity.fields);
    if (!resolvedUrl) {
      return;
    }

    window.open(resolvedUrl, '_blank', 'noopener,noreferrer');
  }, []);

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.container}>

        <div className={styles.header}>
          <Text as="h2" size={500} weight="semibold" className={styles.title}>
            Entity Search
          </Text>
          <Text size={200} className={styles.subtitle}>
            Search and access documents for your entities
          </Text>
        </div>

        <div className={styles.searchWrapper}>
          <SearchBox
            placeholder="Search by name, type, or state..."
            value={searchQuery}
            onChange={(_, data) => {
              setSearchQuery(data.value);
              setCurrentPage(1);
            }}
            size="medium"
            style={{ width: '100%' }}
          />
        </div>

        <div className={styles.resultsHeader}>
          <Text size={200} className={styles.resultsCount}>
            {isLoading ? 'Loading entities...' : `${filtered.length} ${filtered.length === 1 ? 'entity' : 'entities'} found`}
          </Text>
        </div>

        {parseError && (
          <div className={styles.actionConfigError}>
            <Text size={200}>Invalid actions JSON. Fix the web part setting to render row actions.</Text>
          </div>
        )}

        <div className={styles.resultsList}>
          <div className={styles.listHeader}>
            <span>Name</span>
            <span>Actions</span>
          </div>
          {error ? (
            <div className={styles.noResults}>
              <Text size={300} className={styles.noResultsText}>
                {error}
              </Text>
            </div>
          ) : !isLoading && filtered.length === 0 ? (
            <div className={styles.noResults}>
              <Text size={300} className={styles.noResultsText}>
                No entities match &ldquo;{searchQuery}&rdquo;
              </Text>
            </div>
          ) : (
            pagedEntities.map((entity) => (
              <React.Fragment key={entity.id}>
                <div className={styles.entityRow}>
                  <div className={styles.entityInfo}>
                    <div className={styles.entityNameRow}>
                      <Text size={300} weight="semibold">{entity.name}</Text>
                      <Badge
                        appearance="filled"
                        color={entity.status.toLowerCase() === 'active' ? 'success' : 'subtle'}
                        size="small"
                      >
                        {entity.status || 'Unknown'}
                      </Badge>
                    </div>
                    <Text size={200} className={styles.entityMeta}>
                      {entity.type}&nbsp;&middot;&nbsp;{entity.deal}
                    </Text>
                  </div>

                  <div className={styles.entityActions}>
                    {actions.map((action, actionIndex) => {
                      if (action.type === 'single') {
                        return (
                          <Button
                            key={`single-${actionIndex}-${action.label}`}
                            size="small"
                            appearance="subtle"
                            onClick={() => onActionClick(action.url, entity)}
                          >
                            {action.label}
                          </Button>
                        );
                      }

                      const dropdownKey = `${entity.id}-${actionIndex}`;
                      const isDropdownOpen = openMenuKey === dropdownKey;

                      return (
                        <div
                          key={`dropdown-${actionIndex}-${action.label}`}
                          className={styles.dropdownWrapper}
                          onClick={e => e.stopPropagation()}
                        >
                          <Button
                            size="small"
                            appearance="subtle"
                            onClick={() => setOpenMenuKey(isDropdownOpen ? null : dropdownKey)}
                          >
                            {action.label} ▾
                          </Button>
                          {isDropdownOpen && (
                            <div className={styles.dropdownMenu}>
                              {action.items.map(item => (
                                <div
                                  key={item.label}
                                  className={styles.dropdownItem}
                                  onClick={() => {
                                    onActionClick(item.url, entity);
                                    setOpenMenuKey(null);
                                  }}
                                >
                                  {item.label}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </React.Fragment>
            ))
          )}

          {!error && !isLoading && filtered.length > 0 && (
            <div className={styles.paginationBar}>
              <Text size={200} className={styles.paginationInfo}>
                Showing {pageStartNumber}-{pageEndNumber} of {filtered.length}
              </Text>
              <div className={styles.paginationButtons}>
                <Button
                  size="small"
                  appearance="secondary"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((prevPage) => prevPage - 1)}
                >
                  Previous
                </Button>
                <Text size={200} className={styles.paginationInfo}>
                  Page {currentPage} of {totalPages}
                </Text>
                <Button
                  size="small"
                  appearance="secondary"
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage((prevPage) => prevPage + 1)}
                >
                  Next
                </Button>
              </div>
            </div>
          )}
        </div>

      </div>
    </FluentProvider>
  );
};

function readFieldValue(item: Record<string, unknown>, fieldInternalName: string): string {
  if (!fieldInternalName) {
    return '';
  }

  return readUnknownFieldValue(item[fieldInternalName]);
}

function readUnknownFieldValue(rawValue: unknown): string {
  if (rawValue === null || rawValue === undefined) {
    return '';
  }

  if (typeof rawValue === 'string' || typeof rawValue === 'number' || typeof rawValue === 'boolean') {
    return String(rawValue);
  }

  if (typeof rawValue === 'object') {
    const valueObject = rawValue as { Title?: string; LookupValue?: string; Label?: string; Value?: string };
    const simplifiedValue = valueObject.LookupValue || valueObject.Title || valueObject.Label || valueObject.Value;
    return simplifiedValue ? String(simplifiedValue) : '';
  }

  return '';
}

function mapItemFields(item: Record<string, unknown>): Record<string, string> {
  return Object.keys(item).reduce((accumulator, key) => {
    accumulator[key] = readUnknownFieldValue(item[key]);
    return accumulator;
  }, {} as Record<string, string>);
}

function resolveActionUrl(urlTemplate: string, fields: Record<string, string>): string {
  return urlTemplate.replace(/\{\{\s*([A-Za-z0-9_]+)\s*\}\}/g, (_fullMatch, fieldName: string) => {
    return fields[fieldName] || '';
  });
}

function extractTemplateFields(actionsConfigurationJson: string): string[] {
  if (!actionsConfigurationJson) {
    return [];
  }

  const fieldSet = new Set<string>();
  const fieldPattern = /\{\{\s*([A-Za-z0-9_]+)\s*\}\}/g;
  let match: RegExpExecArray | null = fieldPattern.exec(actionsConfigurationJson);

  while (match) {
    fieldSet.add(match[1]);
    match = fieldPattern.exec(actionsConfigurationJson);
  }

  return Array.from(fieldSet);
}

function parseActionsConfiguration(actionsConfigurationJson: string): { actions: RowAction[]; parseError: string } {
  if (!actionsConfigurationJson || !actionsConfigurationJson.trim()) {
    return {
      actions: [],
      parseError: ''
    };
  }

  try {
    const parsedValue = JSON.parse(actionsConfigurationJson) as unknown;
    if (!Array.isArray(parsedValue)) {
      return {
        actions: [],
        parseError: 'Actions JSON root must be an array.'
      };
    }

    const actions: RowAction[] = [];

    parsedValue.forEach((candidateAction) => {
      if (!candidateAction || typeof candidateAction !== 'object') {
        return;
      }

      const action = candidateAction as {
        type?: unknown;
        label?: unknown;
        url?: unknown;
        items?: Array<{ label?: unknown; url?: unknown }>;
      };

      if (action.type !== 'single' && action.type !== 'dropdown') {
        return;
      }

      if (typeof action.label !== 'string' || !action.label.trim()) {
        return;
      }

      if (action.type === 'single' && typeof action.url === 'string' && action.url.trim()) {
        actions.push({
          type: 'single',
          label: action.label.trim(),
          url: action.url
        });
      }

      if (action.type === 'dropdown' && Array.isArray(action.items)) {
        const items = action.items
          .filter((item) => item && typeof item.label === 'string' && item.label.trim() && typeof item.url === 'string' && item.url.trim())
          .map((item) => ({
            label: String(item.label).trim(),
            url: String(item.url)
          }));

        if (items.length > 0) {
          actions.push({
            type: 'dropdown',
            label: action.label.trim(),
            items
          });
        }
      }
    });

    return {
      actions,
      parseError: ''
    };
  } catch (parseError) {
    return {
      actions: [],
      parseError: parseError instanceof Error ? parseError.message : 'Unable to parse actions JSON.'
    };
  }
}

export default EntitySearchWebpart;
