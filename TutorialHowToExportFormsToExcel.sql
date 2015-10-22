--Esse tutorial explicará como criar um arquivo do Excel, a partir de um cursor, ou de 
--vários campos, conforme a necessidade, usando OLE2. 



--Primeiramente, crie a seguinte procedure PREENCHE_CEL: 

--CÓDIGO: SELECIONAR TODOS

-- ESSA PROCEDURE É A RESPONSÁVEL PELO PREENCHIMENTO DE CADA CÉLULA DO ARQUIVO DO EXCEL.
PROCEDURE PREENCHE_CEL(WORKSHEET IN OUT OLE2.OBJ_TYPE, 
                       CELL IN OUT OLE2.OBJ_TYPE,
                       ARGS IN OUT OLE2.LIST_TYPE, 
                       Row_num number,                    -- linha
                       Col_num number,                    -- coluna
                       TITULO VARCHAR2,                   -- o que vai ser inserido na célula
                       COL_WIDTH NUMBER,                  -- tamanho da coluna
                       FONT_NAME VARCHAR2,                -- nome da fonte
                       FONT_SIZE VARCHAR2,                -- tamanho da fonte
                       FONT_BOLD BOOLEAN,                 -- deve ser bold?
                       FONT_ITAL BOOLEAN,                 -- deve ser itálico?
                       COR_INDEX NUMBER,                  -- índice da cor da fonte
                       Align binary_integer DEFAULT null, -- alinhamento horizontal do texto
                       formato VARCHAR2 DEFAULT NULL,    
                       -- formato de entrada do dado ('Geral', '0','#.##0,00', 'dd/mm/aa', 'd/m/aa h:mm AM/PM') 
                       
                       Tipo varchar2 default 'VALUE',     -- tipo do dado, se valor ('VALUE') ou fórumla ('FORMULA')
                       BGCOR_INDEX NUMBER DEFAULT 0)IS    -- índice da cor de fundo
FONT       OLE2.OBJ_TYPE;
v_interior OLE2.OBJ_TYPE;
  
BEGIN
  ARGS := OLE2.CREATE_ARGLIST;
   OLE2.ADD_ARG(ARGS, Row_num); -- ROW NUMBER
   OLE2.ADD_ARG(ARGS, Col_num); -- COLUMN NUMBER
   CELL := OLE2.GET_OBJ_PROPERTY(WORKSHEET, 'CELLS', ARGS);
   OLE2.DESTROY_ARGLIST(ARGS);
   OLE2.SET_PROPERTY(CELL, Tipo, TITULO); 
   if COL_WIDTH is not null then
      OLE2.SET_PROPERTY(CELL, 'COLUMNWIDTH', COL_WIDTH);
   end if; 
   font := ole2.get_obj_property (cell, 'Font');
  OLE2.SET_PROPERTY (font, 'Name', FONT_NAME);
  OLE2.SET_PROPERTY (font, 'Size', FONT_SIZE);
  OLE2.SET_PROPERTY (font, 'Bold', FONT_BOLD);
  OLE2.SET_PROPERTY (font, 'Italic', FONT_ITAL);
  -- ALTERA CORES DA ÁRVORE 

  OLE2.SET_PROPERTY(font, 'ColorIndex', COR_INDEX);  --0,Preto (3, Red)
  if Align is not null then
     OLE2.SET_PROPERTY(CELL, 'HorizontalAlignment', Align);
  end if; 
  if formato is not null then
     OLE2.SET_PROPERTY(CELL, 'NumberFormat', formato);
  END IF;
  
  v_interior := ole2.get_obj_property(CELL,'Interior');
  ole2.set_property(v_interior,'ColorIndex',BGCOR_INDEX); 
  ole2.release_obj(v_interior);

   OLE2.RELEASE_OBJ(font);   
   OLE2.RELEASE_OBJ(CELL);
END;




--Compile. Não deve dar nenhum erro.
--Agora, declare no botão que vai fazer a exportação, o seguinte:

--CÓDIGO: SELECIONAR TODOS

  --WHEN-BUTTON-PRESSED por exemplo:
     
  DECLARE
     -- DECLARA VARIÁVEIS PARA OS OBJETOS. 
     APPLICATION OLE2.OBJ_TYPE; 
     WORKBOOKS OLE2.OBJ_TYPE; 
     WORKBOOK OLE2.OBJ_TYPE; 
     WORKSHEET OLE2.OBJ_TYPE; 
     CELL OLE2.OBJ_TYPE; 
     FONT OLE2.OBJ_TYPE; 
     
     -- DECLARA RECIPIENTES PARA LISTAS DE ARGUMENTOS OLE 
     ARGS OLE2.LIST_TYPE; 
     V_ALERT      number; 
     ROWCOUNT     NUMBER := 1; -- contador de linhas
     COLCOUNT     NUMBER := 1; -- contador de colunas
     V_NOME       VARCHAR2( 260 ) ; 
     DIRETORIO    VARCHAR2( 256 ) ; 
     V_DIR_MODELO VARCHAR2( 60 ) ; 
     
     -- DECLARA SUBTIPOS DE FORMATAÇÃO
     SUBTYPE xlHAlign IS binary_integer; 
     CENTER                CONSTANT xlHAlign := - 4108; 
     CENTERACROSSSELECTION CONSTANT xlHAlign := 7; 
     DISTRIBUTED           CONSTANT xlHAlign := - 4117; 
     FILL                  CONSTANT xlHAlign := 5; 
     GENERAL               CONSTANT xlHAlign := 1; 
     JUSTIFY               CONSTANT xlHAlign := - 4130; 
     LEFT                  CONSTANT xlHAlign := - 4131; 
     RIGHT                 CONSTANT xlHAlign := - 4152;    
  BEGIN
     
      ...
      
      SET_APPLICATION_PROPERTY( CURSOR_STYLE, 'BUSY' ) ; -- cursor de sistema ocupado
      
      -- DECLARA RECIPIENTES PARA OBJETO DE APLICAÇÃO
     APPLICATION := OLE2.CREATE_OBJ( 'EXCEL.APPLICATION' ) ; 
     -- CRIA UMA COLEÇÃO DE WORKBOOKS E ADICIONA UM NOVO WORKBOOK
     WORKBOOKS := OLE2.GET_OBJ_PROPERTY( APPLICATION, 'WORKBOOKS' ) ; 
     WORKBOOK  := OLE2.GET_OBJ_PROPERTY( WORKBOOKS, 'ADD' ) ; 
     -- ABRE A WORKSHEET PLAN1 NO WORKBOOK
     ARGS := OLE2.CREATE_ARGLIST; 
     OLE2.ADD_ARG( ARGS, 'PLAN1' ) ; 
     WORKSHEET := OLE2.GET_OBJ_PROPERTY( WORKBOOK, 'WORKSHEETS', ARGS ) ; 
     OLE2.DESTROY_ARGLIST( ARGS ) ;
     ...
    /* Parâmetros: 
           PREENCHE_CEL(WORKSHEET,CELL,ARGS,Row_num,Col_num,TITULO CHAR,COL_WIDTH NUMBER,FONT_NAME VARCHAR2,FONT_SIZE VARCHAR2,
           FONT_BOLD BOOLEAN,FONT_ITAL BOOLEAN,COR_INDEX NUMBER, ALINHAMENTO, FORMATO NUMERICO, value ou formula?, BGColor )IS  
           0=Preto; 3=Vermelho; 5=Dark Blue ; 13=Cinza 
    */   
    ...
    -- supondo que C_GERAL seja seu cursor que retorna todos os dados necessários, faça:
     for d in C_GERAL 
     LOOP 
        
        COLCOUNT   := 1; -- DIZ QUE ELE DEVE COMEÇAR A PREENCHER NA PRIMEIRA COLUNA
      
        PREENCHE_CEL( WORKSHEET, CELL, ARGS, ROWCOUNT, COLCOUNT, d.id_prof     , NULL, 'Arial', '10', FALSE, FALSE, 0 ) ; 
        -- codigo do profissional
        
        COLCOUNT := COLCOUNT + 1;  -- AGORA ELE DEVE PREENCHER NA SEGUNDA COLUNA
        PREENCHE_CEL( WORKSHEET, CELL, ARGS, ROWCOUNT, COLCOUNT, d.profissional, NULL, 'Verdana', '10', TRUE, FALSE, 0 ) ; 
        -- nome dele em bold Verdana
        
        COLCOUNT := COLCOUNT + 1;  -- AGORA ELE DEVE PREENCHER NA TERCEIRA COLUNA
        PREENCHE_CEL( WORKSHEET, CELL, ARGS, ROWCOUNT, COLCOUNT, To_char( d.dt_inclusao, 'DD/MM/RRRR' ) , 
        NULL, 'Arial', '10', FALSE, FALSE, 0, NULL, 'dd/mm/aaaa' ) ; 
        
        COLCOUNT := COLCOUNT + 1;  -- AGORA ELE DEVE PREENCHER NA QUARTA COLUNA
        PREENCHE_CEL( WORKSHEET, CELL, ARGS, ROWCOUNT, COLCOUNT, d.devedor , NULL, 'Arial', '10', FALSE, FALSE, 0, 
        NULL, '#.##0,00', 'VALUE', decode(d.devedor,0,3,5) ) ; 
        -- se devedor = 0, o fundo fica vermelho. Se devedor diferente de 0, fundo fica azul.
                
        COLCOUNT := COLCOUNT + 1;  -- AGORA ELE DEVE PREENCHER NA QUINTA COLUNA
        PREENCHE_CEL( WORKSHEET, CELL, ARGS, ROWCOUNT, COLCOUNT, v_p.total     , NULL, 'Arial', '10', FALSE, TRUE, 0, 
        NULL, '#.##0,00' ) ; -- total em itálico
        
        ROWCOUNT := ROWCOUNT + 1; -- ADICIONA UMA LINHA
     END LOOP;
     
     -- PRONTO, ELE CARREGOU TODOS OS DADOS DO CURSOR PARA DENTRO DO ARQUIVO EXCEL, TOTALMENTE FORMATADO.
        
     -- PERMITE AO USER VER A APLICAÇÃO DO EXCEL PARA VER O RESULTADO.
     OLE2.SET_PROPERTY( APPLICATION, 'VISIBLE', TRUE ) ; 
     ---------------------------------------------------------------------------------------------------------------- 
      -- SALVANDO O ARQUIVO 
      
               V_NOME = 'O_NOME_DO_ARQUIVO.XLS'; 
               DIRETORIO    := 'C:\sua_pasta\'||V_NOME; 
               V_DIR_MODELO := 'C:\sua_pasta\'; 
               
               ARGS         := OLE2.CREATE_ARGLIST; 
               OLE2.ADD_ARG( ARGS, DIRETORIO ) ; 
               OLE2.INVOKE( WORKSHEET, 'SaveAs', ARGS ) ; 
               OLE2.DESTROY_ARGLIST( ARGS ) ; 
     ------------------------------------------------------------------------------------------------------------------ 
       --FECHANDO O ARQUIVO E APLICAÇÃO -- comente para não fechar automaticamente.
               /*
               ARGS := OLE2.CREATE_ARGLIST; 
               OLE2.ADD_ARG(ARGS, 0);
               OLE2.INVOKE(WORKBOOK, 'Close', ARGS);
               OLE2.DESTROY_ARGLIST(ARGS);
               --*/ 
      ----------------------------------------------------------------------------------------------------------------  
       -- LIBERA RECIPIENTES DA MEMÓRIA
               OLE2.RELEASE_OBJ( WORKSHEET ) ; 
               OLE2.RELEASE_OBJ( WORKBOOK ) ; 
               OLE2.RELEASE_OBJ( WORKBOOKS ) ; 
               OLE2.RELEASE_OBJ( APPLICATION ) ; 
      
      ----------------------------------------------------------------------------------------------------------------  
       -- EXIBE UMA MENSAGEM CONFIRMANDO               
               
       SET_APPLICATION_PROPERTY( CURSOR_STYLE, 'DEFAULT' ) ; -- cursor volta ao normal.
       SET_ALERT_PROPERTY( 'AVISO', ALERT_MESSAGE_TEXT, 'Planilha gerada com sucesso em '|| DIRETORIO ) ; 
       V_ALERT := SHOW_ALERT( 'AVISO' ) ; 
    
  EXCEPTION          -- CASO ACONTEÇA ALGUMA COISA ERRADA NO MEIO DO CAMINHO:
  WHEN OTHERS THEN 

     SET_APPLICATION_PROPERTY( CURSOR_STYLE, 'DEFAULT' ) ; 
     CLEAR_MESSAGE; 
     OLE2.RELEASE_OBJ( WORKSHEET ) ; 
     OLE2.RELEASE_OBJ( WORKBOOK ) ; 
     OLE2.RELEASE_OBJ( WORKBOOKS ) ; 
     OLE2.Release_Obj( application ) ; 

     message( 'Error'||sqlerrm ) ; 

     SET_ALERT_PROPERTY( 'AVISO', ALERT_MESSAGE_TEXT, 'Erro ao salvar o arquivo' ) ; 
     V_ALERT := SHOW_ALERT( 'AVISO' ) ; 
     RAISE FORM_TRIGGER_FAILURE; 
  END;  
  


--Para escolher as cores, http://www.mvps.org/dmcritchie/excel/colors.htm 
