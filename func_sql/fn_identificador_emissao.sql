CREATE FUNCTION dbo.fn_identificador_emissao
(
    @relatorio_id VARCHAR(6),
    @nome_emissao VARCHAR(255)
)
RETURNS VARCHAR(500)
AS
BEGIN
    RETURN
        FORMAT(
            CONVERT(
                date,
                LEFT(@relatorio_id, 4) + RIGHT(@relatorio_id, 2) + '01'
            ),
            'dd/MM/yyyy'
        )
        + ' - ' +
        @nome_emissao;
END;
GO
