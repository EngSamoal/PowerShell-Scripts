/**
 * db.js — IndexedDB access layer for the Sales & Inventory System.
 *
 * Single source of truth for the database schema. All other modules read
 * and write through the helpers exported here instead of touching
 * indexedDB directly, so the schema only needs to change in one place.
 */

const DB_NAME = 'SalesSystemDB';
const DB_VERSION = 2;

let dbInstance = null;

/**
 * Opens (and if needed creates/upgrades) the database.
 * Safe to call multiple times — returns the same open connection.
 */
function openDatabase() {
  return new Promise((resolve, reject) => {
    if (dbInstance) {
      resolve(dbInstance);
      return;
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = event.target.result;

      if (!db.objectStoreNames.contains('products')) {
        const products = db.createObjectStore('products', { keyPath: 'id', autoIncrement: true });
        products.createIndex('sku', 'sku', { unique: true });
        products.createIndex('barcode', 'barcode', { unique: false });
        products.createIndex('name', 'name', { unique: false });
        products.createIndex('category', 'category', { unique: false });
        products.createIndex('status', 'status', { unique: false });
      }

      if (!db.objectStoreNames.contains('customers')) {
        const customers = db.createObjectStore('customers', { keyPath: 'id', autoIncrement: true });
        customers.createIndex('phone', 'phone', { unique: false });
        customers.createIndex('customerType', 'customerType', { unique: false });
      }

      if (!db.objectStoreNames.contains('invoices')) {
        const invoices = db.createObjectStore('invoices', { keyPath: 'id', autoIncrement: true });
        invoices.createIndex('invoiceNumber', 'invoiceNumber', { unique: true });
        invoices.createIndex('date', 'date', { unique: false });
        invoices.createIndex('customerId', 'customerId', { unique: false });
        invoices.createIndex('paymentStatus', 'paymentStatus', { unique: false });
        invoices.createIndex('status', 'status', { unique: false });
      }

      if (!db.objectStoreNames.contains('invoiceItems')) {
        const invoiceItems = db.createObjectStore('invoiceItems', { keyPath: 'id', autoIncrement: true });
        invoiceItems.createIndex('invoiceId', 'invoiceId', { unique: false });
        invoiceItems.createIndex('productId', 'productId', { unique: false });
      }

      if (!db.objectStoreNames.contains('expenses')) {
        const expenses = db.createObjectStore('expenses', { keyPath: 'id', autoIncrement: true });
        expenses.createIndex('date', 'date', { unique: false });
        expenses.createIndex('expenseType', 'expenseType', { unique: false });
      }

      if (!db.objectStoreNames.contains('stockAdjustments')) {
        const stockAdjustments = db.createObjectStore('stockAdjustments', { keyPath: 'id', autoIncrement: true });
        stockAdjustments.createIndex('productId', 'productId', { unique: false });
        stockAdjustments.createIndex('date', 'date', { unique: false });
      }

      if (!db.objectStoreNames.contains('settings')) {
        db.createObjectStore('settings', { keyPath: 'key' });
      }

      if (!db.objectStoreNames.contains('counters')) {
        db.createObjectStore('counters', { keyPath: 'name' });
      }

      if (!db.objectStoreNames.contains('auditLog')) {
        const auditLog = db.createObjectStore('auditLog', { keyPath: 'id', autoIncrement: true });
        auditLog.createIndex('date', 'date', { unique: false });
        auditLog.createIndex('entityType', 'entityType', { unique: false });
      }

      if (!db.objectStoreNames.contains('returns')) {
        const returns = db.createObjectStore('returns', { keyPath: 'id', autoIncrement: true });
        returns.createIndex('invoiceId', 'invoiceId', { unique: false });
        returns.createIndex('date', 'date', { unique: false });
      }
    };

    request.onsuccess = (event) => {
      dbInstance = event.target.result;
      resolve(dbInstance);
    };

    request.onerror = () => {
      reject(request.error);
    };
  });
}

/** Wraps an IDBRequest in a Promise. */
function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/** Reads a single record by key. */
async function dbGet(storeName, key) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, 'readonly');
  return requestToPromise(tx.objectStore(storeName).get(key));
}

/** Reads every record in a store. */
async function dbGetAll(storeName) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, 'readonly');
  return requestToPromise(tx.objectStore(storeName).getAll());
}

/** Reads records matching a value on a given index. */
async function dbGetAllByIndex(storeName, indexName, value) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, 'readonly');
  return requestToPromise(tx.objectStore(storeName).index(indexName).getAll(value));
}

/** Inserts or updates a record. Returns the record's key. */
async function dbPut(storeName, value) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, 'readwrite');
  const result = await requestToPromise(tx.objectStore(storeName).put(value));
  return result;
}

/** Deletes a record by key. */
async function dbDelete(storeName, key) {
  const db = await openDatabase();
  const tx = db.transaction(storeName, 'readwrite');
  await requestToPromise(tx.objectStore(storeName).delete(key));
}

/**
 * Runs a function inside a single readwrite transaction spanning the given
 * stores, so a multi-store change (e.g. sale = new invoice + stock update)
 * either fully succeeds or fully fails together.
 *
 * `work` receives the transaction and must use tx.objectStore(name) —
 * not the dbGet/dbPut helpers above, which open their own transactions.
 */
async function dbTransaction(storeNames, work) {
  const db = await openDatabase();
  const tx = db.transaction(storeNames, 'readwrite');
  const result = await work(tx);
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve(result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('Transaction aborted'));
  });
}
