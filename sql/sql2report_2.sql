-- sheet_param category_name = Sports | Family | Action
-- chart 
-- section_heading=category_name
SELECT
    *
FROM
    category
WHERE name = :category_name;


-- chart
-- section_heading=Shortest 10 films in category_name
SELECT
    f.film_id, f.title, c.name, f.length
FROM film f
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
WHERE name = :category_name
ORDER BY f.length
LIMIT 10;


-- chart=bar
-- section_heading=Distribution of film length in category_name
-- query_param min_length=60 
-- query_param max_length=90
SELECT
    CONCAT(FLOOR(f.length / 10) * 10, ' - ', FLOOR(f.length / 10) * 10 + 9) AS length_range,
    count(*) as n_films
FROM film f
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
WHERE name = :category_name AND length between :min_length and :max_length
GROUP BY FLOOR(f.length / 10) * 10, length_range
ORDER BY FLOOR(f.length / 10) * 10;

-- chart
-- section_heading=Film in category_name with length in 1-2 hours (using query parameters)
-- query_param min_length=60 
-- query_param max_length=120
SELECT
    f.film_id, f.title, c.name, f.length
FROM film f
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
WHERE name = :category_name AND length between :min_length and :max_length
ORDER BY f.length;

-- chart=bar
-- section_heading=Top Rented Films in category_name for rating rating_p
-- query_param rating_p=PG | PG-13 | R
SELECT 
    f.title,
    COUNT(DISTINCT r.rental_id) AS total_rentals
FROM film f
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
    JOIN inventory i ON f.film_id = i.film_id
    JOIN rental r ON i.inventory_id = r.inventory_id
    JOIN payment p ON r.rental_id = p.rental_id
WHERE name = :category_name AND f.rating = :rating_p
GROUP BY f.film_id, f.title
ORDER BY total_rentals DESC
LIMIT 10;