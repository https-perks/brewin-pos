CREATE TABLE items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sku TEXT UNIQUE,
  name TEXT NOT NULL,
  category TEXT,
  price REAL NOT NULL DEFAULT 0,
  active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE components (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT UNIQUE NOT NULL,
  qty_used REAL DEFAULT 0,
  unit_cost REAL DEFAULT 0,
  pos_track_sellout INTEGER DEFAULT 0
);

CREATE TABLE recipes (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  item_id INTEGER NOT NULL,
  component_id INTEGER NOT NULL,
  qty_per_item REAL NOT NULL DEFAULT 0,
  FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE,
  FOREIGN KEY (component_id) REFERENCES components(id) ON DELETE CASCADE,
  UNIQUE (item_id, component_id)
);

CREATE TABLE inventory (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT UNIQUE NOT NULL,
  unit TEXT,
  case_cost REAL DEFAULT 0,
  units_per_case REAL DEFAULT 0,
  qty_on_hand REAL DEFAULT 0,
  reorder_point REAL DEFAULT 0
);

INSERT INTO items (id, sku, name, category, price, active) VALUES (1, 'COF-8', 'Coffee 8oz', 'Drinks', 1.5, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (2, 'COF-12', 'Coffee 12oz', 'Food', 2.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (3, 'TEA', 'Hot Tea', 'Drinks', 3.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (4, 'HOT-CHOC', 'Hot Cocoa', 'Drinks', 3.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (5, 'PAS-SM', 'Sm. Pastry', 'Food', 2.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (6, 'PAS-LG', 'Lg. Pastry', 'Food', 3.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (7, 'SNA-CHEX', 'Chex Mix', 'Food', 2.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (8, 'SNA-JERK', 'Jerky Stick', 'Food', 2.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (9, 'SNA-BELV', 'Belvita', 'Food', 2.0, 1);
INSERT INTO items (id, sku, name, category, price, active) VALUES (10, 'SNA-COOK', 'Cookie', 'Food', 2.0, 1);

INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (1, 'Small Pastry', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (2, 'Foam Cups 8oz', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (3, 'Foam Cups 12oz', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (4, 'Large Pastry', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (5, 'Coffee Stirrers', 0.0, 0.0, 0);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (6, 'Belvita Crackers', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (7, 'Chex Mix', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (8, 'Jerky Stick', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (9, 'Cookie', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (10, 'Compostable Plates', 0.0, 0.0, 0);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (11, 'Lids 12oz', 0.0, 0.0, 0);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (12, 'Lids 8oz', 0.0, 0.0, 0);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (13, 'Bigelow Tea', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (14, 'Hot Chocolate', 0.0, 0.0, 1);
INSERT INTO components (id, name, qty_used, unit_cost, pos_track_sellout) VALUES (15, 'Napkins', 0.0, 0.0, 0);

INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (1, 9, 6, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (2, 7, 7, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (3, 10, 9, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (4, 8, 8, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (5, 1, 2, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (6, 1, 12, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (7, 1, 5, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (8, 2, 3, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (9, 2, 11, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (10, 2, 5, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (11, 4, 2, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (12, 4, 12, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (13, 4, 5, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (14, 4, 14, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (15, 3, 13, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (16, 3, 3, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (17, 3, 11, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (18, 3, 5, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (19, 6, 10, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (20, 6, 4, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (21, 5, 1, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (22, 5, 10, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (23, 6, 15, 1.0);
INSERT INTO recipes (id, item_id, component_id, qty_per_item) VALUES (24, 5, 15, 1.0);

INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (1, 'Small Pastry', 'Each', 0.75, 1.0, 0.0, 0.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (2, 'Large Pastry', 'Each', 1.0, 1.0, 0.0, 0.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (5, 'Coffee Filters', 'Each', 33.0, 1249.0, 1000.0, 50.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (6, 'Folger''s Coffee', 'Each', 35.0, 36.0, 36.0, 50.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (7, 'N''Joy Creamer (12oz)', 'Each', 10.25, 3.0, 3.0, 1.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (8, 'N''Joy Sugar (20oz)', 'Each', 11.0, 3.0, 3.0, 1.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (9, 'Compostable Plates', 'Each', 75.0, 1000.0, 1000.0, 100.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (10, 'Foam Cups 12oz', 'Each', 70.0, 1000.0, 998.0, 100.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (11, 'Foam Cups 8oz', 'Each', 55.0, 1000.0, 1000.0, 100.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (12, 'Coffee Stirrers', 'Each', 4.3, 1000.0, 1000.0, 100.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (13, 'Napkins', 'Each', 92.0, 500.0, 500.0, 100.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (14, 'Bigelow Tea', 'Each', 44.0, 168.0, 167.0, 28.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (15, 'Hot Chocolate', 'Each', 18.0, 50.0, 0.0, 25.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (16, 'Chex Mix', 'Each', 38.71, 60.0, 0.0, 10.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (17, 'Jerky Stick', 'Each', 20.75, 20.0, 0.0, 10.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (18, 'Belvita Crackers', 'Each', 8.9, 8.0, 0.0, 10.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (19, 'Cookie', 'Each', 10.35, 12.0, 0.0, 10.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (22, 'Lids 12oz', 'Each', 70.0, 1000.0, 998.0, 100.0);
INSERT INTO inventory (id, name, unit, case_cost, units_per_case, qty_on_hand, reorder_point) VALUES (23, 'Lids 8oz', 'Each', 55.0, 1000.0, 1000.0, 100.0);

