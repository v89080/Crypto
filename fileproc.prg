PROCEDURE WRITE             && Вывод строки лицевого счета
      * Вызывается из процедуры ExcelAccount (см. выше)
      oExcel.Range([A]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=DTOC(pAccount.Data_zap)

      oExcel.Range([B]+ALLTRIM(STR(nRow,3))).Select
      DO CASE
         CASE pAccount.CONTENTS='Сальдо старт'
              cAcContents=[Начальное сальдо по договору ]
         CASE pAccount.CONTENTS='Упл.Аренда  '
              cAcContents=[Уплачено за аренду]       
         CASE pAccount.CONTENTS='Упл.НДС '
              cAcContents=[Уплачено НДС]       
         CASE pAccount.CONTENTS='Упл.Пени    '
              cAcContents=[Уплачено пени]       
         CASE pAccount.CONTENTS='Упл.Штраф   '
              cAcContents=[Уплачено штрафа]       
         CASE pAccount.CONTENTS='Возв.Аренда '
              cAcContents=[Возвращено переплат за аренду]       
         CASE pAccount.CONTENTS='Возв.НДС    '
              cAcContents=[Возвращено переплат НДС]       
         CASE pAccount.CONTENTS='Возв.Пени   '
              cAcContents=[Возвращено пени]       
         CASE pAccount.CONTENTS='Спис.Аренда '
              cAcContents=[Списано за аренду (Арбитраж)]       
         CASE pAccount.CONTENTS='Спис.НДС    '
              cAcContents=[Списано НДС (Арбитраж)]       
         CASE pAccount.CONTENTS='Спис.Пени   '
              cAcContents=[Списано пени (Арбитраж)]
         CASE pAccount.CONTENTS='НачисленоАр '
              cAcContents=[Начислено к уплате за аренду]       
         CASE pAccount.CONTENTS='НачисленоНДС'
              cAcContents=[Начислено к уплате НДС]       
         CASE pAccount.CONTENTS='УменьшеноАр '
              cAcContents=[Уменьшено за аренду по льготам ]       
         CASE pAccount.CONTENTS='УменьшеноНДС'
              cAcContents=[Уменьшено НДС по льготам ]       
         CASE pAccount.CONTENTS='Списано пени'
              cAcContents=[Списано пени по льготам]       
         CASE pAccount.CONTENTS='ДоначислАр  '
              cAcContents=[Доначислено к уплате за аренду]       
         CASE pAccount.CONTENTS='ДоначислНДС '
              cAcContents=[Доначислено к уплате НДС]       
         CASE pAccount.CONTENTS='Штраф       '
              cAcContents=[Начислен штраф]       
         CASE pAccount.CONTENTS='ИзмСтавкиРеф'
              cAcContents=[Изменение ставки рефинансирования]       
         CASE pAccount.CONTENTS='Сальдо      '
              cAcContents=[Сальдо]       
         CASE pAccount.CONTENTS='Конец догов.'
              cAcContents=[Окончание срока действия договора]       
         OTHERWISE
             cAcContents=pAccount.Contents
      ENDCASE   
      oExcel.ActiveCell.FormulaR1C1=cAcContents
      oExcel.Range([C]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.nach
      oExcel.Range([D]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.ymyb
      oExcel.Range([E]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.ypdona
      oExcel.Range([F]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.ypvs
      oExcel.Range([G]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.vpe
      oExcel.Range([H]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.saldo
      oExcel.Range([I]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.npe
      oExcel.Range([J]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.remark
      oExcel.Range([K]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.ypnedo
      oExcel.Range([L]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.raypvo
      oExcel.Range([M]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.ostpe
      oExcel.Range([N]+ALLTRIM(STR(nRow,3))).Select
      oExcel.ActiveCell.FormulaR1C1=pAccount.ostsh
RETURN
