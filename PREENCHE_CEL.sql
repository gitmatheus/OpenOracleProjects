
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


