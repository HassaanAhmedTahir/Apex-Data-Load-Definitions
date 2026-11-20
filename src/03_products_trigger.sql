CREATE OR REPLACE TRIGGER prc_products_bu
BEFORE UPDATE ON prc_products
FOR EACH ROW
BEGIN
  :NEW.prod_updated    := SYSTIMESTAMP;
  :NEW.prod_updated_by := COALESCE(SYS_CONTEXT('apex$session','app_user'),
                                   SYS_CONTEXT('userenv','session_user'));
END;
/
