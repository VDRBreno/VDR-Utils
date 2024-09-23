CREATE OR REPLACE FUNCTION get_common_prefix_length(str1 TEXT, str2 TEXT)
	RETURNS float8 as $$
DECLARE
	min_length integer := LEAST(LENGTH(str1), LENGTH(str2));
	i integer;
BEGIN
	FOR i IN 0..min_length LOOP
		IF NOT SUBSTRING(str1 FROM i FOR 1) = SUBSTRING(str2 FROM i FOR 1) THEN
			RETURN i;
		END IF;
	END LOOP;

	RETURN min_length;
END;
$$
LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION jaro_winkler(str1 TEXT, str2 TEXT)
	RETURNS float8 as $$
DECLARE
	len1 integer := LENGTH(str1);
	len2 integer := LENGTH(str2);
	match_distance integer;
	match1 bool[];
	match2 bool[];
	matches integer := 0;
	transp integer := 0;
	i integer;
	j integer;
	k integer;
	l integer;
	great integer;
	low integer;
	jaro_sim float8;
	common_prefix_len float8;
	jaro_winkler_sim float8;
BEGIN
	FOR i IN 0..len1 LOOP
		match1[i] := false;
	END LOOP;
	FOR i IN 0..len2 LOOP
		match2[i] := false;
	END LOOP;

	IF len1 = 0 OR len2 = 0 THEN
		RETURN 0;
	END IF;

	match_distance := FLOOR(GREATEST(len1, len2) / 2) - 1;
	
	FOR i IN 0..len1 LOOP
		great := GREATEST(0, i - match_distance);
		low := LEAST(i + match_distance + 1, len2);
		<<inner>>
		FOR j IN great..low LOOP
			IF NOT match2[j] AND SUBSTRING(str1 FROM i FOR 1) = SUBSTRING(str2 FROM j FOR 1) THEN
				match1[i] := true;
				match2[j] := true;
				matches := matches + 1;
				EXIT inner;
			END IF;
		END LOOP;
	END LOOP;

	IF matches = 0 THEN
		RETURN 0;
	END IF;

	k := 0;
	FOR l IN 0..len1 LOOP
		IF match1[l] THEN
			WHILE NOT match2[k] LOOP
				k := k + 1;
			END LOOP;

			IF NOT SUBSTRING(str1 FROM l FOR 1) = SUBSTRING(str2 FROM k FOR 1) THEN
				transp := transp + 1;
			END IF;

			k := k + 1;
		END IF;
	END LOOP;

	jaro_sim := (matches / len1 + matches / len2 + (matches - transp / 2) / matches) / 3;
	common_prefix_len := get_common_prefix_length(str1, str2);
	jaro_winkler_sim := jaro_sim + 0.1 * common_prefix_len * (1 - jaro_sim);

	RETURN jaro_winkler_sim;	
END;
$$
LANGUAGE plpgsql;