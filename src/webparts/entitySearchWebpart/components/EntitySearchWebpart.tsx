import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  SearchBox,
  Button,
  Badge,
  Text,
} from '@fluentui/react-components';
import styles from './EntitySearchWebpart.module.scss';
import type { IEntitySearchWebpartProps } from './IEntitySearchWebpartProps';

interface IEntity {
  id: number;
  name: string;
  type: string;
  deal  : string;
  status: 'Active' | 'Inactive';
}

const SAMPLE_ENTITIES: IEntity[] = [
  { id: 1, name: 'Acme Holdings LLC',           type: 'LLC',         deal: 'Deal 1', status: 'Active'   },
  { id: 2, name: 'Acme Company', type: 'Partnership', deal: 'Deal 2', status: 'Active'   },
];

const MENU_ITEMS = ['Archive', 'Board Minutes', 'State Documents', 'Tax & Accounting'];

const EntitySearchWebpart: React.FC<IEntitySearchWebpartProps> = () => {
  const [searchQuery, setSearchQuery] = React.useState('');
  const [openMenuId, setOpenMenuId] = React.useState<number | null>(null);

  const filtered = SAMPLE_ENTITIES.filter(e =>
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
            {filtered.length} {filtered.length === 1 ? 'entity' : 'entities'} found
          </Text>
        </div>

        <div className={styles.resultsList}>
          <div className={styles.listHeader}>
            <span>Name</span>
            <span>Actions</span>
          </div>
          {filtered.length === 0 ? (
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
                        color={entity.status === 'Active' ? 'success' : 'subtle'}
                        size="small"
                      >
                        {entity.status}
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

export default EntitySearchWebpart;
