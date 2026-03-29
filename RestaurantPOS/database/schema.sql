-- ============================================================
-- Restaurant POS SaaS Platform - SQL Server Database Schema
-- ============================================================

USE master;
GO

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'RestaurantPOS')
    CREATE DATABASE RestaurantPOS;
GO

USE RestaurantPOS;
GO

-- ============================================================
-- TENANTS / LOCATIONS
-- ============================================================
CREATE TABLE tenants (
    tenant_id       INT IDENTITY(1,1) PRIMARY KEY,
    company_name    NVARCHAR(200) NOT NULL,
    slug            NVARCHAR(100) NOT NULL UNIQUE,
    email           NVARCHAR(200) NOT NULL,
    phone           NVARCHAR(30),
    address         NVARCHAR(500),
    logo_url        NVARCHAR(500),
    plan            NVARCHAR(50) NOT NULL DEFAULT 'starter', -- starter, professional, enterprise
    is_active       BIT NOT NULL DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE(),
    updated_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE locations (
    location_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    location_name   NVARCHAR(200) NOT NULL,
    address         NVARCHAR(500),
    city            NVARCHAR(100),
    state           NVARCHAR(100),
    zip             NVARCHAR(20),
    country         NVARCHAR(100) DEFAULT 'US',
    phone           NVARCHAR(30),
    email           NVARCHAR(200),
    timezone        NVARCHAR(100) DEFAULT 'UTC',
    currency        NVARCHAR(10) DEFAULT 'USD',
    tax_rate        DECIMAL(5,4) DEFAULT 0.0800,
    is_active       BIT NOT NULL DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- USERS & ROLES
-- ============================================================
CREATE TABLE roles (
    role_id         INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    role_name       NVARCHAR(100) NOT NULL,
    permissions     NVARCHAR(MAX), -- JSON array of permission keys
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE users (
    user_id         INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    location_id     INT REFERENCES locations(location_id),
    role_id         INT REFERENCES roles(role_id),
    first_name      NVARCHAR(100) NOT NULL,
    last_name       NVARCHAR(100) NOT NULL,
    email           NVARCHAR(200) NOT NULL,
    phone           NVARCHAR(30),
    password_hash   NVARCHAR(500) NOT NULL,
    pin_code        NVARCHAR(10), -- for quick POS login
    avatar_url      NVARCHAR(500),
    is_active       BIT NOT NULL DEFAULT 1,
    last_login      DATETIME2,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE(),
    CONSTRAINT UQ_users_email_tenant UNIQUE (email, tenant_id)
);

CREATE TABLE sessions (
    session_id      NVARCHAR(128) PRIMARY KEY,
    user_id         INT NOT NULL REFERENCES users(user_id),
    ip_address      NVARCHAR(45),
    user_agent      NVARCHAR(500),
    payload         NVARCHAR(MAX),
    last_activity   INT NOT NULL,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- MENU & PRODUCTS
-- ============================================================
CREATE TABLE menu_categories (
    category_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    parent_id       INT REFERENCES menu_categories(category_id),
    category_name   NVARCHAR(200) NOT NULL,
    description     NVARCHAR(1000),
    image_url       NVARCHAR(500),
    sort_order      INT DEFAULT 0,
    is_active       BIT NOT NULL DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE menu_items (
    item_id         INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    category_id     INT REFERENCES menu_categories(category_id),
    item_name       NVARCHAR(200) NOT NULL,
    description     NVARCHAR(2000),
    price           DECIMAL(10,2) NOT NULL,
    cost_price      DECIMAL(10,2) DEFAULT 0,
    sku             NVARCHAR(100),
    barcode         NVARCHAR(100),
    image_url       NVARCHAR(500),
    item_type       NVARCHAR(50) DEFAULT 'food', -- food, beverage, combo, modifier
    calories        INT,
    prep_time_min   INT DEFAULT 0,
    is_available    BIT NOT NULL DEFAULT 1,
    is_taxable      BIT NOT NULL DEFAULT 1,
    track_inventory BIT NOT NULL DEFAULT 0,
    sort_order      INT DEFAULT 0,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE(),
    updated_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE item_variants (
    variant_id      INT IDENTITY(1,1) PRIMARY KEY,
    item_id         INT NOT NULL REFERENCES menu_items(item_id),
    variant_name    NVARCHAR(200) NOT NULL,
    price_modifier  DECIMAL(10,2) DEFAULT 0,
    sku             NVARCHAR(100),
    is_default      BIT DEFAULT 0
);

CREATE TABLE modifier_groups (
    group_id        INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    group_name      NVARCHAR(200) NOT NULL,
    selection_type  NVARCHAR(20) DEFAULT 'single', -- single, multiple
    min_select      INT DEFAULT 0,
    max_select      INT DEFAULT 1,
    is_required     BIT DEFAULT 0
);

CREATE TABLE modifiers (
    modifier_id     INT IDENTITY(1,1) PRIMARY KEY,
    group_id        INT NOT NULL REFERENCES modifier_groups(group_id),
    modifier_name   NVARCHAR(200) NOT NULL,
    price_add       DECIMAL(10,2) DEFAULT 0,
    is_available    BIT DEFAULT 1,
    sort_order      INT DEFAULT 0
);

CREATE TABLE item_modifier_groups (
    item_id         INT NOT NULL REFERENCES menu_items(item_id),
    group_id        INT NOT NULL REFERENCES modifier_groups(group_id),
    PRIMARY KEY (item_id, group_id)
);

-- ============================================================
-- TABLES / FLOOR PLAN
-- ============================================================
CREATE TABLE dining_areas (
    area_id         INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    area_name       NVARCHAR(100) NOT NULL,
    description     NVARCHAR(500),
    sort_order      INT DEFAULT 0
);

CREATE TABLE restaurant_tables (
    table_id        INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    area_id         INT REFERENCES dining_areas(area_id),
    table_number    NVARCHAR(20) NOT NULL,
    capacity        INT DEFAULT 4,
    qr_code_token   NVARCHAR(100) UNIQUE,
    status          NVARCHAR(20) DEFAULT 'available', -- available, occupied, reserved, cleaning
    pos_x           INT DEFAULT 0,
    pos_y           INT DEFAULT 0,
    shape           NVARCHAR(20) DEFAULT 'rectangle',
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- ORDERS
-- ============================================================
CREATE TABLE orders (
    order_id        INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    table_id        INT REFERENCES restaurant_tables(table_id),
    customer_id     INT,  -- FK added after customers table
    user_id         INT REFERENCES users(user_id),
    order_number    NVARCHAR(50) NOT NULL,
    order_type      NVARCHAR(30) NOT NULL DEFAULT 'dine-in', -- dine-in, takeout, delivery, qr-order, kiosk
    status          NVARCHAR(30) NOT NULL DEFAULT 'pending', -- pending, confirmed, preparing, ready, served, completed, cancelled
    subtotal        DECIMAL(10,2) NOT NULL DEFAULT 0,
    tax_amount      DECIMAL(10,2) NOT NULL DEFAULT 0,
    discount_amount DECIMAL(10,2) NOT NULL DEFAULT 0,
    tip_amount      DECIMAL(10,2) NOT NULL DEFAULT 0,
    delivery_fee    DECIMAL(10,2) NOT NULL DEFAULT 0,
    total_amount    DECIMAL(10,2) NOT NULL DEFAULT 0,
    notes           NVARCHAR(1000),
    source          NVARCHAR(50) DEFAULT 'pos', -- pos, qr, kiosk, delivery_app, online
    scheduled_at    DATETIME2,
    completed_at    DATETIME2,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE(),
    updated_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE order_items (
    order_item_id   INT IDENTITY(1,1) PRIMARY KEY,
    order_id        INT NOT NULL REFERENCES orders(order_id),
    item_id         INT NOT NULL REFERENCES menu_items(item_id),
    variant_id      INT REFERENCES item_variants(variant_id),
    quantity        INT NOT NULL DEFAULT 1,
    unit_price      DECIMAL(10,2) NOT NULL,
    modifier_total  DECIMAL(10,2) DEFAULT 0,
    line_total      DECIMAL(10,2) NOT NULL,
    notes           NVARCHAR(500),
    status          NVARCHAR(30) DEFAULT 'pending', -- pending, preparing, ready, served, void
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE order_item_modifiers (
    id              INT IDENTITY(1,1) PRIMARY KEY,
    order_item_id   INT NOT NULL REFERENCES order_items(order_item_id),
    modifier_id     INT NOT NULL REFERENCES modifiers(modifier_id),
    modifier_name   NVARCHAR(200),
    price_add       DECIMAL(10,2) DEFAULT 0
);

-- ============================================================
-- CUSTOMERS / CRM
-- ============================================================
CREATE TABLE customers (
    customer_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    first_name      NVARCHAR(100),
    last_name       NVARCHAR(100),
    email           NVARCHAR(200),
    phone           NVARCHAR(30),
    date_of_birth   DATE,
    gender          NVARCHAR(10),
    address         NVARCHAR(500),
    city            NVARCHAR(100),
    notes           NVARCHAR(1000),
    loyalty_points  INT DEFAULT 0,
    total_visits    INT DEFAULT 0,
    total_spent     DECIMAL(12,2) DEFAULT 0,
    last_visit      DATETIME2,
    segment         NVARCHAR(50) DEFAULT 'regular', -- new, regular, vip, at-risk, lost
    is_active       BIT DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE(),
    updated_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

ALTER TABLE orders ADD CONSTRAINT FK_orders_customer FOREIGN KEY (customer_id) REFERENCES customers(customer_id);

CREATE TABLE loyalty_programs (
    program_id      INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    program_name    NVARCHAR(200) NOT NULL,
    points_per_dollar DECIMAL(5,2) DEFAULT 1,
    redeem_rate     DECIMAL(5,4) DEFAULT 0.01, -- 1 point = $0.01
    min_redeem      INT DEFAULT 100,
    is_active       BIT DEFAULT 1
);

CREATE TABLE loyalty_transactions (
    txn_id          INT IDENTITY(1,1) PRIMARY KEY,
    customer_id     INT NOT NULL REFERENCES customers(customer_id),
    order_id        INT REFERENCES orders(order_id),
    points_earned   INT DEFAULT 0,
    points_redeemed INT DEFAULT 0,
    balance_after   INT NOT NULL,
    txn_type        NVARCHAR(30), -- earn, redeem, adjustment, expire
    notes           NVARCHAR(500),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE customer_feedback (
    feedback_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    customer_id     INT REFERENCES customers(customer_id),
    order_id        INT REFERENCES orders(order_id),
    rating          TINYINT CHECK (rating BETWEEN 1 AND 5),
    food_rating     TINYINT,
    service_rating  TINYINT,
    ambiance_rating TINYINT,
    comment         NVARCHAR(2000),
    source          NVARCHAR(50) DEFAULT 'qr', -- qr, kiosk, email, manual
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- INVENTORY
-- ============================================================
CREATE TABLE inventory_categories (
    inv_cat_id      INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    cat_name        NVARCHAR(200) NOT NULL,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE inventory_items (
    inv_item_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    inv_cat_id      INT REFERENCES inventory_categories(inv_cat_id),
    item_name       NVARCHAR(200) NOT NULL,
    sku             NVARCHAR(100),
    unit            NVARCHAR(50) NOT NULL, -- kg, liter, piece, box
    quantity_on_hand DECIMAL(10,3) DEFAULT 0,
    reorder_level   DECIMAL(10,3) DEFAULT 0,
    reorder_qty     DECIMAL(10,3) DEFAULT 0,
    cost_per_unit   DECIMAL(10,4) DEFAULT 0,
    supplier_id     INT,
    last_restocked  DATETIME2,
    expiry_date     DATE,
    is_active       BIT DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE(),
    updated_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE suppliers (
    supplier_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    supplier_name   NVARCHAR(200) NOT NULL,
    contact_name    NVARCHAR(200),
    email           NVARCHAR(200),
    phone           NVARCHAR(30),
    address         NVARCHAR(500),
    payment_terms   NVARCHAR(100),
    is_active       BIT DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

ALTER TABLE inventory_items ADD CONSTRAINT FK_inv_supplier FOREIGN KEY (supplier_id) REFERENCES suppliers(supplier_id);

CREATE TABLE purchase_orders (
    po_id           INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    supplier_id     INT REFERENCES suppliers(supplier_id),
    po_number       NVARCHAR(50) NOT NULL,
    status          NVARCHAR(30) DEFAULT 'draft', -- draft, sent, received, partial, cancelled
    order_date      DATE NOT NULL DEFAULT CAST(GETDATE() AS DATE),
    expected_date   DATE,
    received_date   DATE,
    subtotal        DECIMAL(12,2) DEFAULT 0,
    tax_amount      DECIMAL(12,2) DEFAULT 0,
    total_amount    DECIMAL(12,2) DEFAULT 0,
    notes           NVARCHAR(1000),
    created_by      INT REFERENCES users(user_id),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE purchase_order_items (
    po_item_id      INT IDENTITY(1,1) PRIMARY KEY,
    po_id           INT NOT NULL REFERENCES purchase_orders(po_id),
    inv_item_id     INT NOT NULL REFERENCES inventory_items(inv_item_id),
    qty_ordered     DECIMAL(10,3) NOT NULL,
    qty_received    DECIMAL(10,3) DEFAULT 0,
    unit_cost       DECIMAL(10,4) NOT NULL,
    line_total      DECIMAL(12,2) NOT NULL
);

CREATE TABLE inventory_movements (
    movement_id     INT IDENTITY(1,1) PRIMARY KEY,
    inv_item_id     INT NOT NULL REFERENCES inventory_items(inv_item_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    movement_type   NVARCHAR(30) NOT NULL, -- in, out, adjustment, waste, transfer
    quantity        DECIMAL(10,3) NOT NULL,
    balance_after   DECIMAL(10,3) NOT NULL,
    reference_id    INT, -- order_id or po_id
    reference_type  NVARCHAR(50),
    notes           NVARCHAR(500),
    created_by      INT REFERENCES users(user_id),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE recipes (
    recipe_id       INT IDENTITY(1,1) PRIMARY KEY,
    item_id         INT NOT NULL REFERENCES menu_items(item_id),
    inv_item_id     INT NOT NULL REFERENCES inventory_items(inv_item_id),
    quantity_used   DECIMAL(10,4) NOT NULL,
    unit            NVARCHAR(50)
);

-- ============================================================
-- PAYMENTS
-- ============================================================
CREATE TABLE payments (
    payment_id      INT IDENTITY(1,1) PRIMARY KEY,
    order_id        INT NOT NULL REFERENCES orders(order_id),
    payment_method  NVARCHAR(50) NOT NULL, -- cash, card, qr_code, wallet, split
    amount          DECIMAL(10,2) NOT NULL,
    tip_amount      DECIMAL(10,2) DEFAULT 0,
    status          NVARCHAR(30) DEFAULT 'pending', -- pending, completed, failed, refunded
    reference_no    NVARCHAR(200),
    gateway         NVARCHAR(100), -- stripe, square, paypal, etc.
    gateway_txn_id  NVARCHAR(200),
    processed_by    INT REFERENCES users(user_id),
    processed_at    DATETIME2,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE refunds (
    refund_id       INT IDENTITY(1,1) PRIMARY KEY,
    payment_id      INT NOT NULL REFERENCES payments(payment_id),
    order_id        INT NOT NULL REFERENCES orders(order_id),
    amount          DECIMAL(10,2) NOT NULL,
    reason          NVARCHAR(500),
    status          NVARCHAR(30) DEFAULT 'pending',
    processed_by    INT REFERENCES users(user_id),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- DISCOUNTS & PROMOTIONS
-- ============================================================
CREATE TABLE discounts (
    discount_id     INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    discount_name   NVARCHAR(200) NOT NULL,
    discount_type   NVARCHAR(30) NOT NULL, -- percentage, fixed, bogo, combo
    value           DECIMAL(10,2) NOT NULL,
    code            NVARCHAR(50),
    min_order_amount DECIMAL(10,2) DEFAULT 0,
    max_uses        INT,
    uses_count      INT DEFAULT 0,
    start_date      DATETIME2,
    end_date        DATETIME2,
    applies_to      NVARCHAR(30) DEFAULT 'order', -- order, item, category
    is_active       BIT DEFAULT 1,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- DELIVERY
-- ============================================================
CREATE TABLE delivery_zones (
    zone_id         INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    zone_name       NVARCHAR(200) NOT NULL,
    min_order       DECIMAL(10,2) DEFAULT 0,
    delivery_fee    DECIMAL(10,2) DEFAULT 0,
    free_delivery_above DECIMAL(10,2),
    estimated_time  INT DEFAULT 30, -- minutes
    is_active       BIT DEFAULT 1
);

CREATE TABLE delivery_orders (
    delivery_id     INT IDENTITY(1,1) PRIMARY KEY,
    order_id        INT NOT NULL REFERENCES orders(order_id),
    zone_id         INT REFERENCES delivery_zones(zone_id),
    driver_id       INT REFERENCES users(user_id),
    platform        NVARCHAR(100), -- in-house, ubereats, doordash, grubhub
    platform_order_id NVARCHAR(200),
    delivery_address NVARCHAR(500),
    customer_lat    DECIMAL(10,7),
    customer_lng    DECIMAL(10,7),
    status          NVARCHAR(30) DEFAULT 'pending', -- pending, assigned, picked_up, in_transit, delivered, failed
    estimated_at    DATETIME2,
    picked_up_at    DATETIME2,
    delivered_at    DATETIME2,
    driver_notes    NVARCHAR(500),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- STAFF MANAGEMENT
-- ============================================================
CREATE TABLE shifts (
    shift_id        INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    user_id         INT NOT NULL REFERENCES users(user_id),
    shift_date      DATE NOT NULL,
    start_time      TIME NOT NULL,
    end_time        TIME NOT NULL,
    break_minutes   INT DEFAULT 0,
    status          NVARCHAR(30) DEFAULT 'scheduled', -- scheduled, started, completed, absent, swapped
    notes           NVARCHAR(500),
    created_by      INT REFERENCES users(user_id),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE time_clocks (
    clock_id        INT IDENTITY(1,1) PRIMARY KEY,
    user_id         INT NOT NULL REFERENCES users(user_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    clock_in        DATETIME2 NOT NULL,
    clock_out       DATETIME2,
    break_start     DATETIME2,
    break_end       DATETIME2,
    total_hours     DECIMAL(5,2),
    overtime_hours  DECIMAL(5,2) DEFAULT 0,
    notes           NVARCHAR(500),
    approved_by     INT REFERENCES users(user_id)
);

CREATE TABLE payroll (
    payroll_id      INT IDENTITY(1,1) PRIMARY KEY,
    user_id         INT NOT NULL REFERENCES users(user_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    period_start    DATE NOT NULL,
    period_end      DATE NOT NULL,
    regular_hours   DECIMAL(6,2) DEFAULT 0,
    overtime_hours  DECIMAL(6,2) DEFAULT 0,
    hourly_rate     DECIMAL(8,2) DEFAULT 0,
    tips_amount     DECIMAL(10,2) DEFAULT 0,
    gross_pay       DECIMAL(10,2) DEFAULT 0,
    deductions      DECIMAL(10,2) DEFAULT 0,
    net_pay         DECIMAL(10,2) DEFAULT 0,
    status          NVARCHAR(30) DEFAULT 'pending', -- pending, approved, paid
    processed_at    DATETIME2,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- FINANCIAL
-- ============================================================
CREATE TABLE cash_drawers (
    drawer_id       INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    user_id         INT NOT NULL REFERENCES users(user_id),
    opened_at       DATETIME2 NOT NULL DEFAULT GETDATE(),
    closed_at       DATETIME2,
    opening_float   DECIMAL(10,2) NOT NULL DEFAULT 0,
    closing_amount  DECIMAL(10,2),
    expected_amount DECIMAL(10,2),
    variance        DECIMAL(10,2),
    status          NVARCHAR(20) DEFAULT 'open' -- open, closed
);

CREATE TABLE expenses (
    expense_id      INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    location_id     INT NOT NULL REFERENCES locations(location_id),
    category        NVARCHAR(100) NOT NULL,
    description     NVARCHAR(500),
    amount          DECIMAL(10,2) NOT NULL,
    expense_date    DATE NOT NULL,
    receipt_url     NVARCHAR(500),
    approved_by     INT REFERENCES users(user_id),
    created_by      INT REFERENCES users(user_id),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE tax_rates (
    tax_id          INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    tax_name        NVARCHAR(100) NOT NULL,
    rate            DECIMAL(5,4) NOT NULL,
    applies_to      NVARCHAR(50) DEFAULT 'all', -- all, food, beverage
    is_inclusive    BIT DEFAULT 0,
    is_active       BIT DEFAULT 1
);

-- ============================================================
-- RESERVATIONS
-- ============================================================
CREATE TABLE reservations (
    reservation_id  INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    customer_id     INT REFERENCES customers(customer_id),
    table_id        INT REFERENCES restaurant_tables(table_id),
    party_size      INT NOT NULL,
    reservation_date DATE NOT NULL,
    reservation_time TIME NOT NULL,
    duration_min    INT DEFAULT 90,
    status          NVARCHAR(30) DEFAULT 'confirmed', -- confirmed, seated, completed, no-show, cancelled
    notes           NVARCHAR(500),
    confirmation_code NVARCHAR(20),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- KDS - Kitchen Display System
-- ============================================================
CREATE TABLE kds_stations (
    station_id      INT IDENTITY(1,1) PRIMARY KEY,
    location_id     INT NOT NULL REFERENCES locations(location_id),
    station_name    NVARCHAR(100) NOT NULL,
    category_ids    NVARCHAR(500), -- comma-separated category_ids
    is_active       BIT DEFAULT 1
);

CREATE TABLE kds_tickets (
    ticket_id       INT IDENTITY(1,1) PRIMARY KEY,
    order_id        INT NOT NULL REFERENCES orders(order_id),
    station_id      INT NOT NULL REFERENCES kds_stations(station_id),
    status          NVARCHAR(30) DEFAULT 'new', -- new, in_progress, ready, recalled
    bumped_at       DATETIME2,
    recalled_at     DATETIME2,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- NOTIFICATIONS & AUDIT
-- ============================================================
CREATE TABLE notifications (
    notif_id        INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL REFERENCES tenants(tenant_id),
    user_id         INT REFERENCES users(user_id),
    type            NVARCHAR(50) NOT NULL, -- low_stock, new_order, payment, shift, feedback
    title           NVARCHAR(200) NOT NULL,
    message         NVARCHAR(1000),
    is_read         BIT DEFAULT 0,
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

CREATE TABLE audit_logs (
    log_id          INT IDENTITY(1,1) PRIMARY KEY,
    tenant_id       INT NOT NULL,
    user_id         INT,
    action          NVARCHAR(100) NOT NULL,
    table_name      NVARCHAR(100),
    record_id       INT,
    old_values      NVARCHAR(MAX),
    new_values      NVARCHAR(MAX),
    ip_address      NVARCHAR(45),
    created_at      DATETIME2 NOT NULL DEFAULT GETDATE()
);

-- ============================================================
-- INDEXES FOR PERFORMANCE
-- ============================================================
CREATE INDEX IX_orders_tenant_location ON orders(tenant_id, location_id);
CREATE INDEX IX_orders_status ON orders(status);
CREATE INDEX IX_orders_created ON orders(created_at);
CREATE INDEX IX_order_items_order ON order_items(order_id);
CREATE INDEX IX_inventory_location ON inventory_items(location_id);
CREATE INDEX IX_customers_tenant ON customers(tenant_id);
CREATE INDEX IX_users_tenant ON users(tenant_id);
CREATE INDEX IX_payments_order ON payments(order_id);
CREATE INDEX IX_audit_tenant ON audit_logs(tenant_id);
CREATE INDEX IX_notifications_user ON notifications(user_id, is_read);
GO
