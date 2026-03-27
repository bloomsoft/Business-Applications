-- ============================================================
-- Restaurant POS SaaS Platform - SQLite Schema
-- ============================================================
PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS tenants (
    tenant_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name    TEXT NOT NULL,
    slug            TEXT NOT NULL UNIQUE,
    email           TEXT NOT NULL,
    phone           TEXT,
    address         TEXT,
    logo_url        TEXT,
    plan            TEXT NOT NULL DEFAULT 'starter',
    is_active       INTEGER NOT NULL DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS locations (
    location_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    location_name   TEXT NOT NULL,
    address         TEXT,
    city            TEXT,
    state           TEXT,
    zip             TEXT,
    country         TEXT DEFAULT 'US',
    phone           TEXT,
    email           TEXT,
    timezone        TEXT DEFAULT 'UTC',
    currency        TEXT DEFAULT 'USD',
    tax_rate        REAL DEFAULT 0.0800,
    is_active       INTEGER NOT NULL DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS roles (
    role_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    role_name       TEXT NOT NULL,
    permissions     TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS users (
    user_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    location_id     INTEGER REFERENCES locations(location_id),
    role_id         INTEGER REFERENCES roles(role_id),
    first_name      TEXT NOT NULL,
    last_name       TEXT NOT NULL,
    email           TEXT NOT NULL,
    phone           TEXT,
    password_hash   TEXT NOT NULL,
    pin_code        TEXT,
    avatar_url      TEXT,
    is_active       INTEGER NOT NULL DEFAULT 1,
    last_login      TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (email, tenant_id)
);

CREATE TABLE IF NOT EXISTS sessions (
    session_id      TEXT PRIMARY KEY,
    user_id         INTEGER NOT NULL REFERENCES users(user_id),
    ip_address      TEXT,
    user_agent      TEXT,
    payload         TEXT,
    last_activity   INTEGER NOT NULL,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS menu_categories (
    category_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    parent_id       INTEGER REFERENCES menu_categories(category_id),
    category_name   TEXT NOT NULL,
    description     TEXT,
    image_url       TEXT,
    sort_order      INTEGER DEFAULT 0,
    is_active       INTEGER NOT NULL DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS menu_items (
    item_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    category_id     INTEGER REFERENCES menu_categories(category_id),
    item_name       TEXT NOT NULL,
    description     TEXT,
    price           REAL NOT NULL,
    cost_price      REAL DEFAULT 0,
    sku             TEXT,
    barcode         TEXT,
    image_url       TEXT,
    item_type       TEXT DEFAULT 'food',
    calories        INTEGER,
    prep_time_min   INTEGER DEFAULT 0,
    is_available    INTEGER NOT NULL DEFAULT 1,
    is_taxable      INTEGER NOT NULL DEFAULT 1,
    track_inventory INTEGER NOT NULL DEFAULT 0,
    sort_order      INTEGER DEFAULT 0,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS item_variants (
    variant_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    item_id         INTEGER NOT NULL REFERENCES menu_items(item_id),
    variant_name    TEXT NOT NULL,
    price_modifier  REAL DEFAULT 0,
    sku             TEXT,
    is_default      INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS modifier_groups (
    group_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    group_name      TEXT NOT NULL,
    selection_type  TEXT DEFAULT 'single',
    min_select      INTEGER DEFAULT 0,
    max_select      INTEGER DEFAULT 1,
    is_required     INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS modifiers (
    modifier_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    group_id        INTEGER NOT NULL REFERENCES modifier_groups(group_id),
    modifier_name   TEXT NOT NULL,
    price_add       REAL DEFAULT 0,
    is_available    INTEGER DEFAULT 1,
    sort_order      INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS item_modifier_groups (
    item_id         INTEGER NOT NULL REFERENCES menu_items(item_id),
    group_id        INTEGER NOT NULL REFERENCES modifier_groups(group_id),
    PRIMARY KEY (item_id, group_id)
);

CREATE TABLE IF NOT EXISTS dining_areas (
    area_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    area_name       TEXT NOT NULL,
    description     TEXT,
    sort_order      INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS restaurant_tables (
    table_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    area_id         INTEGER REFERENCES dining_areas(area_id),
    table_number    TEXT NOT NULL,
    capacity        INTEGER DEFAULT 4,
    qr_code_token   TEXT UNIQUE,
    status          TEXT DEFAULT 'available',
    pos_x           INTEGER DEFAULT 0,
    pos_y           INTEGER DEFAULT 0,
    shape           TEXT DEFAULT 'rectangle',
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS customers (
    customer_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    first_name      TEXT,
    last_name       TEXT,
    email           TEXT,
    phone           TEXT,
    date_of_birth   TEXT,
    gender          TEXT,
    address         TEXT,
    city            TEXT,
    notes           TEXT,
    loyalty_points  INTEGER DEFAULT 0,
    total_visits    INTEGER DEFAULT 0,
    total_spent     REAL DEFAULT 0,
    last_visit      TEXT,
    segment         TEXT DEFAULT 'regular',
    is_active       INTEGER DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS orders (
    order_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    table_id        INTEGER REFERENCES restaurant_tables(table_id),
    customer_id     INTEGER REFERENCES customers(customer_id),
    user_id         INTEGER REFERENCES users(user_id),
    order_number    TEXT NOT NULL,
    order_type      TEXT NOT NULL DEFAULT 'dine-in',
    status          TEXT NOT NULL DEFAULT 'pending',
    subtotal        REAL NOT NULL DEFAULT 0,
    tax_amount      REAL NOT NULL DEFAULT 0,
    discount_amount REAL NOT NULL DEFAULT 0,
    tip_amount      REAL NOT NULL DEFAULT 0,
    delivery_fee    REAL NOT NULL DEFAULT 0,
    total_amount    REAL NOT NULL DEFAULT 0,
    notes           TEXT,
    source          TEXT DEFAULT 'pos',
    scheduled_at    TEXT,
    completed_at    TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS order_items (
    order_item_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id        INTEGER NOT NULL REFERENCES orders(order_id),
    item_id         INTEGER NOT NULL REFERENCES menu_items(item_id),
    variant_id      INTEGER REFERENCES item_variants(variant_id),
    quantity        INTEGER NOT NULL DEFAULT 1,
    unit_price      REAL NOT NULL,
    modifier_total  REAL DEFAULT 0,
    line_total      REAL NOT NULL,
    notes           TEXT,
    status          TEXT DEFAULT 'pending',
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS order_item_modifiers (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    order_item_id   INTEGER NOT NULL REFERENCES order_items(order_item_id),
    modifier_id     INTEGER NOT NULL REFERENCES modifiers(modifier_id),
    modifier_name   TEXT,
    price_add       REAL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS loyalty_programs (
    program_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id         INTEGER NOT NULL REFERENCES tenants(tenant_id),
    program_name      TEXT NOT NULL,
    points_per_dollar REAL DEFAULT 1,
    redeem_rate       REAL DEFAULT 0.01,
    min_redeem        INTEGER DEFAULT 100,
    is_active         INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS loyalty_transactions (
    txn_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id     INTEGER NOT NULL REFERENCES customers(customer_id),
    order_id        INTEGER REFERENCES orders(order_id),
    points_earned   INTEGER DEFAULT 0,
    points_redeemed INTEGER DEFAULT 0,
    balance_after   INTEGER NOT NULL,
    txn_type        TEXT,
    notes           TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS customer_feedback (
    feedback_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    customer_id     INTEGER REFERENCES customers(customer_id),
    order_id        INTEGER REFERENCES orders(order_id),
    rating          INTEGER CHECK (rating BETWEEN 1 AND 5),
    food_rating     INTEGER,
    service_rating  INTEGER,
    ambiance_rating INTEGER,
    comment         TEXT,
    source          TEXT DEFAULT 'qr',
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS inventory_categories (
    inv_cat_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    cat_name        TEXT NOT NULL,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS suppliers (
    supplier_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    supplier_name   TEXT NOT NULL,
    contact_name    TEXT,
    email           TEXT,
    phone           TEXT,
    address         TEXT,
    payment_terms   TEXT,
    is_active       INTEGER DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS inventory_items (
    inv_item_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id        INTEGER NOT NULL REFERENCES tenants(tenant_id),
    location_id      INTEGER NOT NULL REFERENCES locations(location_id),
    inv_cat_id       INTEGER REFERENCES inventory_categories(inv_cat_id),
    item_name        TEXT NOT NULL,
    sku              TEXT,
    unit             TEXT NOT NULL,
    quantity_on_hand REAL DEFAULT 0,
    reorder_level    REAL DEFAULT 0,
    reorder_qty      REAL DEFAULT 0,
    cost_per_unit    REAL DEFAULT 0,
    supplier_id      INTEGER REFERENCES suppliers(supplier_id),
    last_restocked   TEXT,
    expiry_date      TEXT,
    is_active        INTEGER DEFAULT 1,
    created_at       TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at       TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS purchase_orders (
    po_id           INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    supplier_id     INTEGER REFERENCES suppliers(supplier_id),
    po_number       TEXT NOT NULL,
    status          TEXT DEFAULT 'draft',
    order_date      TEXT NOT NULL DEFAULT (date('now')),
    expected_date   TEXT,
    received_date   TEXT,
    subtotal        REAL DEFAULT 0,
    tax_amount      REAL DEFAULT 0,
    total_amount    REAL DEFAULT 0,
    notes           TEXT,
    created_by      INTEGER REFERENCES users(user_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS purchase_order_items (
    po_item_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    po_id           INTEGER NOT NULL REFERENCES purchase_orders(po_id),
    inv_item_id     INTEGER NOT NULL REFERENCES inventory_items(inv_item_id),
    qty_ordered     REAL NOT NULL,
    qty_received    REAL DEFAULT 0,
    unit_cost       REAL NOT NULL,
    line_total      REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS inventory_movements (
    movement_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    inv_item_id     INTEGER NOT NULL REFERENCES inventory_items(inv_item_id),
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    movement_type   TEXT NOT NULL,
    quantity        REAL NOT NULL,
    balance_after   REAL NOT NULL,
    reference_id    INTEGER,
    reference_type  TEXT,
    notes           TEXT,
    created_by      INTEGER REFERENCES users(user_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS recipes (
    recipe_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    item_id         INTEGER NOT NULL REFERENCES menu_items(item_id),
    inv_item_id     INTEGER NOT NULL REFERENCES inventory_items(inv_item_id),
    quantity_used   REAL NOT NULL,
    unit            TEXT
);

CREATE TABLE IF NOT EXISTS payments (
    payment_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id        INTEGER NOT NULL REFERENCES orders(order_id),
    payment_method  TEXT NOT NULL,
    amount          REAL NOT NULL,
    tip_amount      REAL DEFAULT 0,
    status          TEXT DEFAULT 'pending',
    reference_no    TEXT,
    gateway         TEXT,
    gateway_txn_id  TEXT,
    processed_by    INTEGER REFERENCES users(user_id),
    processed_at    TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS refunds (
    refund_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    payment_id      INTEGER NOT NULL REFERENCES payments(payment_id),
    order_id        INTEGER NOT NULL REFERENCES orders(order_id),
    amount          REAL NOT NULL,
    reason          TEXT,
    status          TEXT DEFAULT 'pending',
    processed_by    INTEGER REFERENCES users(user_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS discounts (
    discount_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id        INTEGER NOT NULL REFERENCES tenants(tenant_id),
    discount_name    TEXT NOT NULL,
    discount_type    TEXT NOT NULL,
    value            REAL NOT NULL,
    code             TEXT,
    min_order_amount REAL DEFAULT 0,
    max_uses         INTEGER,
    uses_count       INTEGER DEFAULT 0,
    start_date       TEXT,
    end_date         TEXT,
    applies_to       TEXT DEFAULT 'order',
    is_active        INTEGER DEFAULT 1,
    created_at       TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS delivery_zones (
    zone_id              INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id          INTEGER NOT NULL REFERENCES locations(location_id),
    zone_name            TEXT NOT NULL,
    min_order            REAL DEFAULT 0,
    delivery_fee         REAL DEFAULT 0,
    free_delivery_above  REAL,
    estimated_time       INTEGER DEFAULT 30,
    is_active            INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS delivery_orders (
    delivery_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id          INTEGER NOT NULL REFERENCES orders(order_id),
    zone_id           INTEGER REFERENCES delivery_zones(zone_id),
    driver_id         INTEGER REFERENCES users(user_id),
    platform          TEXT,
    platform_order_id TEXT,
    delivery_address  TEXT,
    customer_lat      REAL,
    customer_lng      REAL,
    status            TEXT DEFAULT 'pending',
    estimated_at      TEXT,
    picked_up_at      TEXT,
    delivered_at      TEXT,
    driver_notes      TEXT,
    created_at        TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS shifts (
    shift_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    user_id         INTEGER NOT NULL REFERENCES users(user_id),
    shift_date      TEXT NOT NULL,
    start_time      TEXT NOT NULL,
    end_time        TEXT NOT NULL,
    break_minutes   INTEGER DEFAULT 0,
    status          TEXT DEFAULT 'scheduled',
    notes           TEXT,
    created_by      INTEGER REFERENCES users(user_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS time_clocks (
    clock_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER NOT NULL REFERENCES users(user_id),
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    clock_in        TEXT NOT NULL,
    clock_out       TEXT,
    break_start     TEXT,
    break_end       TEXT,
    total_hours     REAL,
    overtime_hours  REAL DEFAULT 0,
    notes           TEXT,
    approved_by     INTEGER REFERENCES users(user_id)
);

CREATE TABLE IF NOT EXISTS payroll (
    payroll_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER NOT NULL REFERENCES users(user_id),
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    period_start    TEXT NOT NULL,
    period_end      TEXT NOT NULL,
    regular_hours   REAL DEFAULT 0,
    overtime_hours  REAL DEFAULT 0,
    hourly_rate     REAL DEFAULT 0,
    tips_amount     REAL DEFAULT 0,
    gross_pay       REAL DEFAULT 0,
    deductions      REAL DEFAULT 0,
    net_pay         REAL DEFAULT 0,
    status          TEXT DEFAULT 'pending',
    processed_at    TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS cash_drawers (
    drawer_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    user_id         INTEGER NOT NULL REFERENCES users(user_id),
    opened_at       TEXT NOT NULL DEFAULT (datetime('now')),
    closed_at       TEXT,
    opening_float   REAL NOT NULL DEFAULT 0,
    closing_amount  REAL,
    expected_amount REAL,
    variance        REAL,
    status          TEXT DEFAULT 'open'
);

CREATE TABLE IF NOT EXISTS expenses (
    expense_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    category        TEXT NOT NULL,
    description     TEXT,
    amount          REAL NOT NULL,
    expense_date    TEXT NOT NULL,
    receipt_url     TEXT,
    approved_by     INTEGER REFERENCES users(user_id),
    created_by      INTEGER REFERENCES users(user_id),
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS tax_rates (
    tax_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    tax_name        TEXT NOT NULL,
    rate            REAL NOT NULL,
    applies_to      TEXT DEFAULT 'all',
    is_inclusive    INTEGER DEFAULT 0,
    is_active       INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS reservations (
    reservation_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id       INTEGER NOT NULL REFERENCES locations(location_id),
    customer_id       INTEGER REFERENCES customers(customer_id),
    table_id          INTEGER REFERENCES restaurant_tables(table_id),
    party_size        INTEGER NOT NULL,
    reservation_date  TEXT NOT NULL,
    reservation_time  TEXT NOT NULL,
    duration_min      INTEGER DEFAULT 90,
    status            TEXT DEFAULT 'confirmed',
    notes             TEXT,
    confirmation_code TEXT,
    created_at        TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS kds_stations (
    station_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    location_id     INTEGER NOT NULL REFERENCES locations(location_id),
    station_name    TEXT NOT NULL,
    category_ids    TEXT,
    is_active       INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS kds_tickets (
    ticket_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id        INTEGER NOT NULL REFERENCES orders(order_id),
    station_id      INTEGER NOT NULL REFERENCES kds_stations(station_id),
    status          TEXT DEFAULT 'new',
    bumped_at       TEXT,
    recalled_at     TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS notifications (
    notif_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL REFERENCES tenants(tenant_id),
    user_id         INTEGER REFERENCES users(user_id),
    type            TEXT NOT NULL,
    title           TEXT NOT NULL,
    message         TEXT,
    is_read         INTEGER DEFAULT 0,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS audit_logs (
    log_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant_id       INTEGER NOT NULL,
    user_id         INTEGER,
    action          TEXT NOT NULL,
    table_name      TEXT,
    record_id       INTEGER,
    old_values      TEXT,
    new_values      TEXT,
    ip_address      TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Indexes
CREATE INDEX IF NOT EXISTS IX_orders_tenant_location ON orders(tenant_id, location_id);
CREATE INDEX IF NOT EXISTS IX_orders_status ON orders(status);
CREATE INDEX IF NOT EXISTS IX_orders_created ON orders(created_at);
CREATE INDEX IF NOT EXISTS IX_order_items_order ON order_items(order_id);
CREATE INDEX IF NOT EXISTS IX_inventory_location ON inventory_items(location_id);
CREATE INDEX IF NOT EXISTS IX_customers_tenant ON customers(tenant_id);
CREATE INDEX IF NOT EXISTS IX_users_tenant ON users(tenant_id);
CREATE INDEX IF NOT EXISTS IX_payments_order ON payments(order_id);
CREATE INDEX IF NOT EXISTS IX_audit_tenant ON audit_logs(tenant_id);
CREATE INDEX IF NOT EXISTS IX_notifications_user ON notifications(user_id, is_read);
