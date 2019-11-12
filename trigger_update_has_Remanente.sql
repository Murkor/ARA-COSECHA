create or replace trigger actHASremPortal before update of SHAPE on REMANENTE_PORTAL
FOR EACH ROW
DECLARE
    calc       number(38,8);
BEGIN
     calc := SDE.ST_AREA(:NEW.SHAPE) / 10000;
     :NEW.HAS := calc;
END;
/