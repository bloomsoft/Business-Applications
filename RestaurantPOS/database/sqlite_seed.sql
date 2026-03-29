-- ============================================================
-- Sample Seed Data for RestaurantPOS (SQLite)
-- ============================================================

-- Tenant
INSERT INTO tenants (company_name, slug, email, phone, plan)
VALUES ('Demo Restaurant Group', 'demo', 'admin@demo.com', '(555) 123-4567', 'professional');

-- Locations
INSERT INTO locations (tenant_id, location_name, address, city, state, zip, phone, timezone, tax_rate)
VALUES (1, 'Downtown Branch', '100 Main Street', 'New York', 'NY', '10001', '(555) 100-0001', 'America/New_York', 0.0875);
INSERT INTO locations (tenant_id, location_name, address, city, state, zip, phone, timezone, tax_rate)
VALUES (1, 'Uptown Branch', '500 Park Avenue', 'New York', 'NY', '10022', '(555) 100-0002', 'America/New_York', 0.0875);

-- Roles
INSERT INTO roles (tenant_id, role_name, permissions) VALUES (1, 'Admin', '["*"]');
INSERT INTO roles (tenant_id, role_name, permissions) VALUES (1, 'Manager', '["pos.access","menu.manage","inventory.view","crm.view","staff.view","financial.view","analytics.view","settings.manage"]');
INSERT INTO roles (tenant_id, role_name, permissions) VALUES (1, 'Cashier', '["pos.access","orders.view"]');
INSERT INTO roles (tenant_id, role_name, permissions) VALUES (1, 'Chef', '["kds.access","orders.view"]');
INSERT INTO roles (tenant_id, role_name, permissions) VALUES (1, 'Waiter', '["pos.access","orders.view","reservations.view"]');
INSERT INTO roles (tenant_id, role_name, permissions) VALUES (1, 'Driver', '["delivery.view"]');

-- Users (password = "password123" hashed with bcrypt cost 12)
INSERT INTO users (tenant_id, location_id, role_id, first_name, last_name, email, phone, password_hash, pin_code)
VALUES (1, 1, 1, 'John', 'Admin', 'admin@demo.com', '(555) 111-0001',
        '$2y$10$CddufnDXtKsOmxQrn7AKwO09BU54WhTDv7hxHLGgX.bxcuXy0OkkW', '1234');
INSERT INTO users (tenant_id, location_id, role_id, first_name, last_name, email, phone, password_hash, pin_code)
VALUES (1, 1, 2, 'Jane', 'Manager', 'manager@demo.com', '(555) 111-0002',
        '$2y$10$CddufnDXtKsOmxQrn7AKwO09BU54WhTDv7hxHLGgX.bxcuXy0OkkW', '5678');
INSERT INTO users (tenant_id, location_id, role_id, first_name, last_name, email, phone, password_hash, pin_code)
VALUES (1, 1, 3, 'Mike', 'Cashier', 'cashier@demo.com', '(555) 111-0003',
        '$2y$10$CddufnDXtKsOmxQrn7AKwO09BU54WhTDv7hxHLGgX.bxcuXy0OkkW', '1111');

-- Menu Categories
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Appetizers', 1);
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Main Courses', 2);
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Burgers', 3);
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Pizza', 4);
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Sides', 5);
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Desserts', 6);
INSERT INTO menu_categories (tenant_id, category_name, sort_order) VALUES (1, 'Beverages', 7);

-- Menu Items
INSERT INTO menu_items (tenant_id, category_id, item_name, description, price, cost_price, item_type, calories, prep_time_min, sort_order) VALUES
(1, 1, 'Caesar Salad', 'Crisp romaine, parmesan, croutons, caesar dressing', 12.99, 3.50, 'food', 350, 8, 1),
(1, 1, 'Buffalo Wings', '8 pcs wings with blue cheese dip', 14.99, 4.00, 'food', 680, 12, 2),
(1, 1, 'Soup of the Day', 'Ask your server for today''s selection', 8.99, 2.00, 'food', 200, 5, 3),
(1, 2, 'Grilled Salmon', 'Atlantic salmon with seasonal vegetables', 24.99, 8.50, 'food', 520, 18, 1),
(1, 2, 'Chicken Parmesan', 'Breaded chicken, marinara, mozzarella, spaghetti', 19.99, 5.50, 'food', 780, 20, 2),
(1, 2, 'Ribeye Steak', '12oz USDA Choice with loaded baked potato', 34.99, 14.00, 'food', 950, 22, 3),
(1, 3, 'Classic Burger', '8oz beef patty, lettuce, tomato, onion, fries', 15.99, 4.50, 'food', 850, 12, 1),
(1, 3, 'BBQ Bacon Burger', 'Cheddar, bacon, onion rings, BBQ sauce, fries', 18.99, 5.50, 'food', 1020, 14, 2),
(1, 3, 'Veggie Burger', 'Plant-based patty, avocado, sprouts, fries', 16.99, 5.00, 'food', 620, 12, 3),
(1, 4, 'Margherita Pizza', 'Fresh mozzarella, basil, tomato sauce', 16.99, 4.00, 'food', 750, 15, 1),
(1, 4, 'Pepperoni Pizza', 'Pepperoni, mozzarella, tomato sauce', 18.99, 4.50, 'food', 880, 15, 2),
(1, 5, 'French Fries', 'Crispy golden fries with ketchup', 5.99, 1.20, 'food', 380, 6, 1),
(1, 5, 'Onion Rings', 'Beer-battered onion rings', 7.99, 1.50, 'food', 420, 8, 2),
(1, 6, 'Chocolate Lava Cake', 'Warm chocolate cake with vanilla ice cream', 10.99, 2.50, 'food', 580, 10, 1),
(1, 6, 'Cheesecake', 'New York style with strawberry topping', 9.99, 2.00, 'food', 450, 3, 2),
(1, 7, 'Soft Drink', 'Coke, Diet Coke, Sprite, or Fanta', 3.49, 0.40, 'beverage', 140, 1, 1),
(1, 7, 'Fresh Juice', 'Orange, apple, or cranberry', 4.99, 1.20, 'beverage', 120, 2, 2),
(1, 7, 'Iced Tea', 'Freshly brewed, sweetened or unsweetened', 3.99, 0.30, 'beverage', 80, 1, 3),
(1, 7, 'Coffee', 'Regular or decaf', 3.49, 0.50, 'beverage', 5, 3, 4),
(1, 7, 'Craft Beer', 'Ask about local selections', 7.99, 2.50, 'beverage', 180, 1, 5);

-- Modifier Groups
INSERT INTO modifier_groups (tenant_id, group_name, selection_type, min_select, max_select, is_required) VALUES
(1, 'Burger Temperature', 'single', 1, 1, 1),
(1, 'Extra Toppings', 'multiple', 0, 5, 0),
(1, 'Pizza Size', 'single', 1, 1, 1),
(1, 'Drink Size', 'single', 1, 1, 1);

-- Modifiers
INSERT INTO modifiers (group_id, modifier_name, price_add, sort_order) VALUES
(1, 'Rare', 0, 1), (1, 'Medium Rare', 0, 2), (1, 'Medium', 0, 3),
(1, 'Medium Well', 0, 4), (1, 'Well Done', 0, 5),
(2, 'Extra Cheese', 1.50, 1), (2, 'Bacon', 2.00, 2), (2, 'Avocado', 2.50, 3),
(2, 'Jalapenos', 1.00, 4), (2, 'Fried Egg', 1.50, 5),
(3, 'Small (10")', -3.00, 1), (3, 'Medium (12")', 0, 2), (3, 'Large (16")', 4.00, 3),
(4, 'Small', 0, 1), (4, 'Medium', 0.50, 2), (4, 'Large', 1.00, 3);

-- Link modifiers to items
INSERT INTO item_modifier_groups (item_id, group_id) VALUES (7, 1), (7, 2), (8, 1), (8, 2), (9, 2);
INSERT INTO item_modifier_groups (item_id, group_id) VALUES (10, 3), (11, 3);
INSERT INTO item_modifier_groups (item_id, group_id) VALUES (16, 4), (17, 4), (18, 4);

-- Dining Areas & Tables
INSERT INTO dining_areas (location_id, area_name, sort_order) VALUES (1, 'Main Hall', 1);
INSERT INTO dining_areas (location_id, area_name, sort_order) VALUES (1, 'Patio', 2);
INSERT INTO dining_areas (location_id, area_name, sort_order) VALUES (1, 'Private Room', 3);

INSERT INTO restaurant_tables (location_id, area_id, table_number, capacity, qr_code_token, pos_x, pos_y) VALUES
(1, 1, '1',  2, 'tok_table_01', 50,  50),
(1, 1, '2',  2, 'tok_table_02', 180, 50),
(1, 1, '3',  4, 'tok_table_03', 310, 50),
(1, 1, '4',  4, 'tok_table_04', 50,  170),
(1, 1, '5',  6, 'tok_table_05', 180, 170),
(1, 1, '6',  6, 'tok_table_06', 310, 170),
(1, 1, '7',  8, 'tok_table_07', 50,  290),
(1, 1, '8',  4, 'tok_table_08', 180, 290),
(1, 2, '9',  4, 'tok_table_09', 440, 50),
(1, 2, '10', 4, 'tok_table_10', 440, 170);

-- Delivery Zones
INSERT INTO delivery_zones (location_id, zone_name, min_order, delivery_fee, free_delivery_above, estimated_time) VALUES
(1, 'Zone A (0-3 miles)', 15.00, 2.99, 35.00, 25),
(1, 'Zone B (3-6 miles)', 20.00, 4.99, 50.00, 40),
(1, 'Zone C (6-10 miles)', 30.00, 7.99, NULL, 55);

-- Loyalty Program
INSERT INTO loyalty_programs (tenant_id, program_name, points_per_dollar, redeem_rate, min_redeem)
VALUES (1, 'Demo Rewards', 1.0, 0.01, 100);

-- Sample Customers
INSERT INTO customers (tenant_id, first_name, last_name, email, phone, loyalty_points, total_visits, total_spent, segment) VALUES
(1, 'Alice', 'Smith',    'alice@example.com',  '(555) 200-0001', 450,  15, 620.50,  'regular'),
(1, 'Bob',   'Johnson',  'bob@example.com',    '(555) 200-0002', 1200, 42, 2150.00, 'vip'),
(1, 'Carol', 'Williams', 'carol@example.com',  '(555) 200-0003', 80,   3,  125.00,  'new'),
(1, 'Dave',  'Brown',    'dave@example.com',   '(555) 200-0004', 320,  8,  410.75,  'regular'),
(1, 'Eve',   'Davis',    'eve@example.com',    '(555) 200-0005', 50,   2,  85.00,   'at-risk');

-- Inventory Categories
INSERT INTO inventory_categories (tenant_id, cat_name) VALUES
(1, 'Proteins'), (1, 'Produce'), (1, 'Dairy'), (1, 'Dry Goods'), (1, 'Beverages'), (1, 'Supplies');

-- Suppliers
INSERT INTO suppliers (tenant_id, supplier_name, contact_name, email, phone, payment_terms) VALUES
(1, 'Fresh Foods Co',  'Tom Fresh',   'orders@freshfoods.com', '(555) 300-0001', 'Net 30'),
(1, 'Metro Beverages', 'Lisa Metro',  'sales@metrobev.com',    '(555) 300-0002', 'Net 15'),
(1, 'Supply Depot',    'Mark Supply', 'info@supplydepot.com',  '(555) 300-0003', 'COD');

-- Inventory Items
INSERT INTO inventory_items (tenant_id, location_id, inv_cat_id, item_name, unit, quantity_on_hand, reorder_level, reorder_qty, cost_per_unit, supplier_id) VALUES
(1, 1, 1, 'Beef Patties (8oz)',   'piece', 120, 30,  60,  2.50,  1),
(1, 1, 1, 'Chicken Breast',       'lb',    45,  15,  30,  4.50,  1),
(1, 1, 1, 'Atlantic Salmon',      'lb',    20,  8,   15,  12.00, 1),
(1, 1, 2, 'Lettuce (Romaine)',    'head',  25,  10,  20,  1.50,  1),
(1, 1, 2, 'Tomatoes',             'lb',    30,  10,  25,  2.00,  1),
(1, 1, 2, 'Onions',               'lb',    40,  15,  30,  1.00,  1),
(1, 1, 3, 'Mozzarella Cheese',    'lb',    18,  8,   15,  5.50,  1),
(1, 1, 3, 'Cheddar Cheese',       'lb',    15,  6,   12,  4.80,  1),
(1, 1, 4, 'Pizza Dough',          'piece', 50,  15,  30,  1.20,  1),
(1, 1, 4, 'Burger Buns',          'piece', 80,  25,  50,  0.60,  1),
(1, 1, 5, 'Coca-Cola (12oz)',     'can',   5,   24,  48,  0.45,  2),
(1, 1, 5, 'Craft Beer (Various)', 'bottle',36,  12,  24,  2.80,  2),
(1, 1, 6, 'To-Go Containers',     'piece', 200, 50,  100, 0.25,  3),
(1, 1, 6, 'Napkins',              'piece', 500, 100, 500, 0.02,  3);

-- KDS Stations
INSERT INTO kds_stations (location_id, station_name, category_ids) VALUES
(1, 'Hot Kitchen',  '2,3,4'),
(1, 'Cold/Prep',    '1,5'),
(1, 'Dessert/Bar',  '6,7');

-- Sample completed orders for analytics
INSERT INTO orders (tenant_id, location_id, table_id, user_id, order_number, order_type, status, subtotal, tax_amount, total_amount, source, completed_at, created_at, updated_at)
VALUES
(1, 1, 1, 1, 'ORD-001', 'dine-in',  'completed', 45.97, 4.02, 49.99, 'pos', datetime('now','-1 hour'),  datetime('now','-2 hours'),  datetime('now','-1 hour')),
(1, 1, 2, 1, 'ORD-002', 'takeout',  'completed', 28.98, 2.54, 31.52, 'pos', datetime('now','-2 hours'), datetime('now','-3 hours'),  datetime('now','-2 hours')),
(1, 1, 3, 2, 'ORD-003', 'delivery', 'completed', 62.97, 5.51, 68.48, 'pos', datetime('now','-3 hours'), datetime('now','-4 hours'),  datetime('now','-3 hours')),
(1, 1, 4, 2, 'ORD-004', 'dine-in',  'completed', 38.98, 3.41, 42.39, 'pos', datetime('now','-4 hours'), datetime('now','-5 hours'),  datetime('now','-4 hours')),
(1, 1, 5, 1, 'ORD-005', 'kiosk',    'completed', 24.97, 2.18, 27.15, 'kiosk', datetime('now','-5 hours'), datetime('now','-6 hours'), datetime('now','-5 hours'));

-- Payments for sample orders
INSERT INTO payments (order_id, payment_method, amount, status, processed_by, processed_at) VALUES
(1, 'card', 49.99, 'completed', 1, datetime('now','-1 hour')),
(2, 'cash', 31.52, 'completed', 1, datetime('now','-2 hours')),
(3, 'card', 68.48, 'completed', 2, datetime('now','-3 hours')),
(4, 'cash', 42.39, 'completed', 2, datetime('now','-4 hours')),
(5, 'card', 27.15, 'completed', 1, datetime('now','-5 hours'));
