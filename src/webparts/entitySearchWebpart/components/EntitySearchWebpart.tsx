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
}

const MENU_ITEMS = ['Archive', 'Board Minutes', 'State Documents', 'Tax & Accounting'];

const EntitySearchWebpart: React.FC<IEntitySearchWebpartProps> = (props) => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [openMenuId, setOpenMenuId] = React.useState<number | null>(null);
  const [entities, setEntities] = React.useState<IEntity[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');

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
        const fieldsToSelect = ['Id', props.titleFieldInternalName, props.typeFieldInternalName, props.dealFieldInternalName, props.statusFieldInternalName]
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
          status: readFieldValue(item, props.statusFieldInternalName)
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

    void loadEntities();

    return () => {
      isActive = false;
    };
  }, [
    props.listId,
    props.titleFieldInternalName,
    props.typeFieldInternalName,
    props.dealFieldInternalName,
    props.statusFieldInternalName,
    props.siteUrl,
    props.spHttpClient
  ]);

  const filtered = entities.filter(e =>
    e.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
    e.type.toLowerCase().includes(searchQuery.toLowerCase()) ||
    e.deal.toLowerCase().includes(searchQuery.toLowerCase())
  );

  // Close menu when clicking outside
  React.useEffect(() => {
    const handleClickOutside = (): void => setOpenMenuId(null);
    document.addEventListener('click', handleClickOutside);
    return () => document.removeEventListener('click', handleClickOutside);
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
            onChange={(_, data) => setSearchQuery(data.value)}
            size="medium"
            style={{ width: '100%' }}
          />
        </div>

        <div className={styles.resultsHeader}>
          <Text size={200} className={styles.resultsCount}>
            {isLoading ? 'Loading entities...' : `${filtered.length} ${filtered.length === 1 ? 'entity' : 'entities'} found`}
          </Text>
        </div>

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
            filtered.map((entity) => (
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
                    <Button size="small" appearance="subtle">
                      Signature Matrix
                    </Button>
                    <div
                      className={styles.dropdownWrapper}
                      onClick={e => e.stopPropagation()}
                    >
                      <Button
                        size="small"
                        appearance="subtle"
                        onClick={() => setOpenMenuId(openMenuId === entity.id ? null : entity.id)}
                      >
                        Documents ▾
                      </Button>
                      {openMenuId === entity.id && (
                        <div className={styles.dropdownMenu}>
                          {MENU_ITEMS.map(item => (
                            <div
                              key={item}
                              className={styles.dropdownItem}
                              onClick={() => setOpenMenuId(null)}
                            >
                              {item}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </React.Fragment>
            ))
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

  const rawValue = item[fieldInternalName];
  if (rawValue === null || rawValue === undefined) {
    return '';
  }

  if (typeof rawValue === 'string' || typeof rawValue === 'number' || typeof rawValue === 'boolean') {
    return String(rawValue);
  }

  if (typeof rawValue === 'object') {
    const valueObject = rawValue as { Title?: string; LookupValue?: string; Label?: string; Value?: string };
    return valueObject.LookupValue || valueObject.Title || valueObject.Label || valueObject.Value || '';
  }

  return '';
}

export default EntitySearchWebpart;
