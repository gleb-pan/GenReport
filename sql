SELECT 
    Timestamp AS 'Дата/Время',
    alarm_message AS 'Сообщение',
    resource AS 'Цех',
    alarm_class,
    log_action 
FROM 
    ALARM_LOG
WHERE 
    Timestamp >= DATEADD(day, -1, CAST(GETDATE() AS DATE)) 
    AND Timestamp < CAST(GETDATE() AS DATE)
ORDER BY
    Timestamp DESC;


SELECT Timestamp AS 'Дата/Время'
, alarm_message AS 'Сообщение'
, resource AS 'Цех'
, alarm_class
, CASE log_action
        WHEN 'G' THEN 'Появление'
        WHEN 'R' THEN 'Исчезновение'
        WHEN 'A' THEN 'Утверждение'
  ELSE 'N/A' END AS 'Статус'
FROM ALARM_LOG
WHERE
    Timestamp >= DATEADD(day, -1, CAST(GETDATE() AS DATE))
    AND Timestamp < CAST(GETDATE() AS DATE)
ORDER BY
    Timestamp DESC;