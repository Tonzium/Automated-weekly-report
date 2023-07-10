#Censored SQL queries

class SQL_queries:
    def __init__(self):
        self.Scrap_SQL_query = self.Scrap_SQL()
        self.Hold_SQL_query = self.Hold_SQL()

    def Scrap_SQL(self):
        Scrap_SQL_query = """
        SELECT 
         AS LOT_ID, 
        ,  
        ,
        ,
        ,
        ,
        CASE WHEN  > 25 THEN 1 ELSE  END AS QTY,
        ,
        ,
        , 
        ,
        TO_CHAR(, ) AS WEEK,
        ,
        REPLACE(, CHR(10), '') AS COMMENTS, 
        
        FROM ,  , 
        WHERE 
             AND
             AND
             AND
             AND
             IS NOT NULL AND
             <> '51' AND
             = '' AND
            DATE BETWEEN TRUNC(SYSDATE-20, 'IW') AND sysdate AND 
            DATE BETWEEN TRUNC(SYSDATE-20, 'IW') AND sysdate AND 
            
        ORDER BY  DESC


        """
        return Scrap_SQL_query

    def Hold_SQL(self):
        Hold_SQL_query = """
        SELECT,  FROM 
        (SELECT  AS LOT_ID,  AS STEP_ID, ,  ,  AS HOLD_DATE,
        REPLACE(, CHR(10), '') AS COMMENTS,
        , TO_CHAR( , ) AS HOLD_WEEK, 
        FROM , , 
        WHERE 
        AND 
        AND 
        AND ( LIKE  OR  = 'ACCEPTANCE_HOLD')
        AND  > 300000000
        AND  IN 
        AND  IS NOT NULL
        AND DATE > SYSDATE-20
        AND DATE > SYSDATE-20
        AND  IN ( , ))
        LEFT OUTER JOIN
        (SELECT  AS STEP_ID,  AS LOCATION_ID
        FROM , 
        WHERE
         AND
         AND
         AND
        
        ON
        """
        return Hold_SQL_query
    
    def get_variables(self):
        return self.Scrap_SQL_query, self.Hold_SQL_query