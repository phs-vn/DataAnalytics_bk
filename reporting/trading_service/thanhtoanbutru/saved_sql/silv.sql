WITH 
[x] AS (
	SELECT [sub_account].[sub_account] FROM (
		SELECT [account_code] FROM [sub_account] [full]
		JOIN [vcf0051] ON [full].[sub_account] = [vcf0051].[sub_account]
		AND [vcf0051].[contract_type] LIKE '%SILV%'
		AND [vcf0051].[date] = '__input__[1]'
	) [acc]
	JOIN [sub_account] ON [sub_account].[account_code] = [acc].[account_code]
),
[i] AS (
	SELECT 
		[relationship].[sub_account],
		[relationship].[account_code],
		[account].[customer_name],
		[branch].[branch_name],
		[c].[approve_date],
		CASE 
			WHEN [c].[approve_date] < '__input__[0]' THEN '__input__[0]'
			ELSE [c].[approve_date]
		END [data_fdate],
		[vcf0051].[contract_type],
		[broker].[broker_name]
	FROM
		[relationship]
	LEFT JOIN
		[vcf0051]
		ON [vcf0051].[sub_account] = [relationship].[sub_account]
		AND [vcf0051].[date] = [relationship].[date]
	LEFT JOIN (
		SELECT 
			[customer_information_change].[account_code], 
			MAX([customer_information_change].[date_of_approval]) [approve_date]
		FROM [customer_information_change]
		WHERE
		    [customer_information_change].[date_of_approval] <= '__input__[1]'
			AND [customer_information_change].[change_content] = 'Loai hinh hop dong'
		GROUP BY
			[customer_information_change].[account_code]
	) [c]
		ON [relationship].[account_code] = [c].[account_code]
	LEFT JOIN
		[account]
		ON [relationship].[account_code] = [account].[account_code]
	LEFT JOIN
		[branch]
		ON [relationship].[branch_id] = [branch].[branch_id]
	LEFT JOIN
		[broker]
		ON [relationship].[broker_id] = [broker].[broker_id]
	RIGHT JOIN 
		[x]
		ON [x].[sub_account] = [relationship].[sub_account]
	WHERE [relationship].[date] = '__input__[1]' 
		AND [vcf0051].[status] IN ('A','B')
		AND DATEDIFF(day,[c].[approve_date],'__input__[1]') > 30
),
[t] AS (
	SELECT 
		[sum].[account_code],
		SUM([sum].[fee])/ROUND(CAST(DATEDIFF(day,[sum].[data_fdate],'__input__[1]')+1 AS FLOAT)/30,0) [fee]
	FROM (
		SELECT 
			[i].[account_code],
			[i].[data_fdate],
			SUM([trading_record].[fee]) [fee]
		FROM
			[trading_record]
		RIGHT JOIN [i] ON [i].[sub_account] = [trading_record].[sub_account]
		WHERE [trading_record].[date] BETWEEN [i].[data_fdate] AND '__input__[1]'
		GROUP BY
			[trading_record].[date],
			[i].[account_code],
			[i].[data_fdate]
	) [sum]
	GROUP BY 
		[sum].[account_code],
		[sum].[data_fdate]
),
[p] AS (
	SELECT
		[sum].[account_code],
		SUM([sum].[fee])/ROUND(CAST(DATEDIFF(day,[sum].[data_fdate],'__input__[1]')+1 AS FLOAT)/30,0) [fee]
	FROM (
		SELECT
			[i].[account_code],
			[i].[data_fdate],
			SUM([payment_in_advance].[total_fee]) [fee]
		FROM
			[payment_in_advance]
		RIGHT JOIN [i] ON [i].[sub_account] = [payment_in_advance].[sub_account]
		WHERE [payment_in_advance].[date] BETWEEN [i].[data_fdate] AND '__input__[1]'
		GROUP BY
			[payment_in_advance].[date],
			[i].[account_code],
			[i].[data_fdate]
	) [sum]
	GROUP BY 
		[sum].[account_code],
		[sum].[data_fdate]
),
[r] AS (
	SELECT
		[sum].[account_code],
		SUM([sum].[int])/ROUND(CAST(DATEDIFF(day,[sum].[data_fdate],'__input__[1]')+1 AS FLOAT)/30,0) [int]
	FROM (
		SELECT
			[i].[account_code],
			[i].[data_fdate],
			SUM([rln0019].[interest]) [int]
		FROM
			[rln0019]
		RIGHT JOIN [i] ON [i].[sub_account] = [rln0019].[sub_account]
		WHERE [rln0019].[date] BETWEEN [i].[data_fdate] AND '__input__[1]'
		GROUP BY
			[rln0019].[date], 
			[i].[account_code],
			[i].[data_fdate]
	) [sum]
	GROUP BY 
		[sum].[account_code],
		[sum].[data_fdate]
),
[silv] AS (
	SELECT [full].* FROM (
		SELECT
			[i].[account_code],
			[i].[customer_name],
			[i].[branch_name],
			[i].[broker_name],
			[i].[approve_date],
			[i].[contract_type],
			ISNULL([t].[fee],0) + 0.3*(ISNULL([p].[fee],0)+ISNULL([r].[int],0)) [fee]
		FROM [i]
		FULL JOIN [t] ON [t].[account_code] = [i].[account_code]
		FULL JOIN [p] ON [p].[account_code] = [i].[account_code]
		FULL JOIN [r] ON [r].[account_code] = [i].[account_code]
	) [full]
),
[anav] AS (
	SELECT 
		[sum].[account_code],
		AVG([sum].[nav]) [nav]
	FROM (
		SELECT 
			[n].[date],
			[i].[account_code],
			SUM([n].[nav]) [nav]
		FROM [i]
		LEFT JOIN (
			SELECT [nav].[date], [nav].[sub_account], [nav].[nav]
			FROM [nav]
			RIGHT JOIN [i] ON [i].[sub_account] = [nav].[sub_account]
			WHERE [nav].[date] BETWEEN [i].[data_fdate] AND '__input__[1]'
		) [n]
		ON [n].[sub_account] = [i].[sub_account]
		GROUP BY 
			[n].[date], 
			[i].[account_code]
	) [sum]
	GROUP BY
		[sum].[account_code]
)
SELECT * FROM (
	SELECT 
		[all].[account_code],
		[all].[customer_name],
		[all].[branch_name],
		[all].[approve_date],
		20000000 [criteria_fee],
		[all].[fee],
		[all].[fee] / 20000000 [pct_fee],
		[all].[nav],
		'SILV PHS' [level],
		CASE 
			WHEN [all].[fee] >= 40000000 THEN 'Promote GOLD PHS'
			WHEN [all].[fee] < 40000000 AND [all].[fee] >= 20000000 THEN 'SILV PHS'
			WHEN [all].[fee] < 20000000 AND [all].[fee] / 20000000 >= 0.8 AND [all].[nav] > 2000000000 THEN 'SILV PHS'
		ELSE
			'Demote DP'
		END [after_review],
		REPLACE([all].[contract_type],' ','') [rate],
		[all].[broker_name]
	FROM (
		SELECT
			[silv].*,
			[anav].[nav]
		FROM [silv]
		LEFT JOIN [anav] ON [anav].[account_code] = [silv].[account_code]
	) [all]
) [final]
WHERE [final].[rate] LIKE '%SILV%'