TRANSFORM Max(L10n_Dict.LngText) AS MaxOfLngText
SELECT
  L10n_Dict.KeyText
FROM
  L10n_Dict
GROUP BY
  L10n_Dict.KeyText PIVOT L10n_Dict.LangCode;
