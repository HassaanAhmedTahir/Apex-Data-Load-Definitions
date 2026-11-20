DROP TRIGGER prc_products_bu;

CREATE OR REPLACE TRIGGER prc_products_biu
BEFORE INSERT OR UPDATE ON prc_products
FOR EACH ROW
BEGIN
  -- normalise text that arrives with odd casing or spaces
  :NEW.prod_status   := UPPER(TRIM(:NEW.prod_status));
  :NEW.prod_currency := REGEXP_REPLACE(UPPER(:NEW.prod_currency), '[^A-Z]', '');

  IF UPDATING THEN
    :NEW.prod_updated    := SYSTIMESTAMP;
    :NEW.prod_updated_by := COALESCE(SYS_CONTEXT('apex$session','app_user'),
                                     SYS_CONTEXT('userenv','session_user'));
  END IF;
END;
/

ALTER TABLE prc_products MODIFY (prod_currency VARCHAR2(10));
