import { join, dirname } from 'path';
import { Low, JSONFile } from 'lowdb';
import { fileURLToPath } from 'url';

const NORTHWIND_DB_DIRECTORY = "northwindDB"; // Directory where LowDB files are stored
const northwindDirectory = join(dirname(dirname(fileURLToPath(import.meta.url))), NORTHWIND_DB_DIRECTORY);

// NOTE: The Northwind database is stored in JSON files using a very simple database
// called lowdb. It does not handle multiple servers, locking, or any guarantee
// of integrity; it's just for development purposes!
//
// To download a fresh copy, type "npm run db-download" from the root of the project.

const tables = {};  // Singleton holds a lowdb object for each "table"
const indexes = {}; // Singleton holds indexes for fast lookup within "tables"

export class dbService {

    async getTable(tableName, primaryKey) {

        const tablesKey = tableName + "::" + primaryKey;
        if (!tables[tablesKey]) {
    
            // If here, there is no table in memory, so read it in now
            const file = join(northwindDirectory, `${tableName}.json`);
            const adapter = new JSONFile(file);
            const db = new Low(adapter);
    
            await db.read();
            db.data = db.data ?? {};
    
            // Store the lowdb and accessors for data and saving changes
            tables[tablesKey] = {
                tableName,
                db,
                primaryKey,
                get data() { return this.db.data[tableName]; },
                item: (value) => {
                    const index = this.getIndex(tables[tablesKey], primaryKey);
                    return index.lookup(value);
                },
                save: () => { this.write(); }
            }
        }
    
        return tables[tablesKey];
    }
    
    getIndex(table, columnName) {
    
        const indexName = `${table.tableName}::${columnName}`;
    
        if (!indexes[indexName]) {
            const index = {};
            for (let i in table.data) {
                index[table.data[i][columnName]] = i;
            }
            indexes[indexName] = {
                index,
                lookup: (key) => table.data[index[key]]
            }
        }
    
        return indexes[indexName];
    }
    
}

