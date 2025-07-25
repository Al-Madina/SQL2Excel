-- chart, section_heading=Writing data without chart
SELECT
    name,
    min(rental_rate),
    avg(rental_rate),
    max(rental_rate)
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
GROUP BY
    1
ORDER BY
    3 DESC;

-- chart=bar, section_heading=Writing data and including a chart
SELECT
    name AS category,
    count(DISTINCT rental_id) AS rental_count
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
    INNER JOIN inventory i ON f.film_id = i.film_id
    INNER JOIN rental r ON i.inventory_id = r.rental_id
GROUP BY
    1
ORDER BY
    2 DESC;

-- Barline is a two-axes plot to present related info in one chart
-- chart=barline, section_heading=Barline chart
SELECT
    ctr.country,
    count(DISTINCT c.customer_id) AS customer_count,
    round(
        100.0 * count(DISTINCT c.customer_id) / (
            SELECT
                count(DISTINCT customer_id)
            FROM
                customer
        ),
        2
    ) AS percentage_share_customer
FROM
    country ctr
    LEFT JOIN city ON ctr.country_id = city.country_id
    LEFT JOIN address a ON city.city_id = a.city_id
    LEFT JOIN customer c ON a.address_id = c.address_id
GROUP BY
    1
ORDER BY
    2 DESC
LIMIT
    10;

-- chart=line, width=25, height=12
-- section_heading=Drawing a chart with some customization, title=Rental count
-- xlabel=Date, ylabel=Rental Count
-- line_width=1.5, line_style=sysDash, marker_symbol=circle, marker_size=8, line_color=A66999
SELECT
    rental_date :: date AS "Rental Date",
    COUNT(DISTINCT rental_id) AS "Rental Count"
FROM
    rental
GROUP BY
    1
ORDER BY
    1;

--chart:bar, data_column_start:3, data_column_end:3
-- section_heading: Selecting consecutive columns using data_column_start and data_column_end
--title: Customer Spending, ylabel: Rental amount, vary_color: True
SELECT
    c.first_name || ' ' || c.last_name AS full_name,
    c.customer_id,
    SUM(p.amount) AS total_amount
FROM
    customer c
    JOIN payment p ON c.customer_id = p.customer_id
WHERE
    c.customer_id IN (1, 2, 3, 4, 5)
GROUP BY
    1,
    2
ORDER BY
    3 DESC;

-- chart:bar
-- data_columns=[2, 4], section_heading=Selecting non-consecutive columns
/* Optionally, you can provide your own headings as below */
--headings=[Film Title, minimum rental rate, Average Rental Rate, maximum rental rate]
SELECT
    name,
    min(rental_rate),
    avg(rental_rate),
    max(rental_rate)
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
GROUP BY
    1
ORDER BY
    3 DESC;

-- chart: pie, column_start=3, width=12, height=12
-- section_heading=Specify the column where the result should be written
-- title=Distribution of films by category
-- show_values=true
SELECT
    c.name AS category,
    COUNT(DISTINCT fc.film_id) AS film_count
FROM
    film_category fc
    JOIN category c ON fc.category_id = c.category_id
GROUP BY
    c.name
ORDER BY
    film_count DESC;

-- The query below should be skipped even though this comment contain the special term chart in it.
SELECT
    name,
    min(rental_rate),
    avg(rental_rate),
    max(rental_rate)
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
GROUP BY
    1
ORDER BY
    3 DESC;

-- chart, section_heading=Query result without pivoting
SELECT
    extract(
        'month'
        FROM
            p.payment_date
    ) || '-' || extract(
        'year'
        FROM
            p.payment_date
    ) AS month_year,
    c.name AS category_name,
    SUM(p.amount) AS total_payment
FROM
    payment p
    JOIN rental r ON p.rental_id = r.rental_id
    JOIN inventory i ON r.inventory_id = i.inventory_id
    JOIN film f ON i.film_id = f.film_id
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
WHERE
    c.name IN ('Family', 'Action', 'Sports', 'Music')
GROUP BY
    1,
    2
ORDER BY
    1,
    2;

-- You can convert data from long-format to wide-format by pivoting the data
-- using index, columns, and values as below
-- chart: bar, section_heading=Query result with pivoting data, chart_position=bottom
-- index=month_year, columns=category_name, values=total_payment
-- show_legend=true, ylabel=Monthly spending, rotation=0
SELECT
    extract(
        'month'
        FROM
            p.payment_date
    ) || '-' || extract(
        'year'
        FROM
            p.payment_date
    ) AS month_year,
    c.name AS category_name,
    SUM(p.amount) AS total_payment
FROM
    payment p
    JOIN rental r ON p.rental_id = r.rental_id
    JOIN inventory i ON r.inventory_id = i.inventory_id
    JOIN film f ON i.film_id = f.film_id
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
WHERE
    c.name IN ('Family', 'Action', 'Sports', 'Music')
GROUP BY
    1,
    2
ORDER BY
    1,
    2;

-- sheetname=my new sheet
-- chart
SELECT
    *
FROM
    category
ORDER BY
    name;

-- This result will not be exported even though the word chart is in the comment
SELECT
    *
FROM
    category
ORDER BY
    name;