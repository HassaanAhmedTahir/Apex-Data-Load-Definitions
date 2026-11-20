-- In SQL Developer or SQL*Plus run SET DEFINE OFF first (the & in "Keyboards & Mice")
INSERT INTO prc_categories (cate_code, cate_name) VALUES ('LAP', 'Laptops');
INSERT INTO prc_categories (cate_code, cate_name) VALUES ('MON', 'Monitors');
INSERT INTO prc_categories (cate_code, cate_name) VALUES ('PER', 'Keyboards & Mice');
INSERT INTO prc_categories (cate_code, cate_name) VALUES ('NET', 'Networking');
INSERT INTO prc_categories (cate_code, cate_name) VALUES ('STO', 'Storage');
INSERT INTO prc_categories (cate_code, cate_name) VALUES ('ACC', 'Accessories');

INSERT INTO prc_suppliers (supp_code, supp_name, supp_country, supp_email)
VALUES ('NORDWIND', 'Nordwind Distribution GmbH', 'DE', 'orders@nordwind.example');
INSERT INTO prc_suppliers (supp_code, supp_name, supp_country, supp_email)
VALUES ('ATLAS', 'Atlas Tech Supply Ltd', 'GB', 'sales@atlas-tech.example');
INSERT INTO prc_suppliers (supp_code, supp_name, supp_country, supp_email)
VALUES ('MERIDIAN', 'Meridian Components B.V.', 'NL', 'sales@meridian.example');
INSERT INTO prc_suppliers (supp_code, supp_name, supp_country, supp_email)
VALUES ('PACRIM', 'Pacific Rim Electronics', 'SG', 'orders@pacrim.example');

COMMIT;
