# Simple program where I can input several movie titles and years and get an Excel Spreedsheet output of more detailed information

#TO DO: Maybe let you search a director and get a spreadsheet of all their movies in this format
# Maybe add a GUI so I can just upload a file instead of having something named movies.txt
# I think maybe instead of country we orient it around spoken language

import os
import sys
import time
import requests
import xlsxwriter
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# CONFIGURATION
# ============================================================

# Put your TMDB API Read Access Token here, or set it as
# an environment variable called TMDB_TOKEN.
TMDB_TOKEN = os.getenv("TMDB_TOKEN")

BASE_URL = "https://api.themoviedb.org/3"

# Files
INPUT_FILE = "movies.txt"
OUTPUT_FILE = "movies.xlsx"

# Language used for TMDB results
LANGUAGE = "en-US"

# ============================================================
# TMDB API
# ============================================================

HEADERS = {
    "Authorization": f"Bearer {TMDB_TOKEN}",
    "accept": "application/json"
}

def search_movie(title, year):
    """
    Search TMDB for a movie using its title and release year.

    Returns the best matching TMDB movie result, or None.
    """

    url = f"{BASE_URL}/search/movie"

    params = {
        "query": title,
        "year": year,
        "language": LANGUAGE,
        "include_adult": "false"
    }

    response = requests.get(
        url,
        headers=HEADERS,
        params=params,
        timeout=30
    )

    response.raise_for_status()

    results = response.json().get("results", [])

    if not results:
        return None

    # Prefer an exact title/year match where possible.
    normalized_title = title.strip().lower()

    exact_matches = []

    for movie in results:
        movie_title = movie.get("title", "").strip().lower()
        release_date = movie.get("release_date", "")

        if (
            movie_title == normalized_title
            and release_date.startswith(str(year))
        ):
            exact_matches.append(movie)

    if exact_matches:
        return exact_matches[0]

    # Otherwise use TMDB's first search result.
    return results[0]


def get_movie_details(movie_id):
    """
    Get detailed movie information and credits in one request.
    """

    url = f"{BASE_URL}/movie/{movie_id}"

    params = {
        "language": LANGUAGE,
        "append_to_response": "credits"
    }

    response = requests.get(
        url,
        headers=HEADERS,
        params=params,
        timeout=30
    )

    response.raise_for_status()

    return response.json()


# ============================================================
# DIRECTOR
# ============================================================

def get_director(details):
    """
    Extract the director(s) from the movie credits.

    Converts names from:
        Christopher Nolan

    To:
        Nolan, Christopher
    """

    crew = details.get("credits", {}).get("crew", [])

    directors = []

    for person in crew:
        if person.get("job") == "Director":

            name = person.get("name", "").strip()

            if not name:
                continue

            # Split the name into individual parts
            name_parts = name.split()

            if len(name_parts) >= 2:
                # Last name goes first
                formatted_name = (
                    f"{name_parts[-1]}, "
                    f"{' '.join(name_parts[:-1])}"
                )
            else:
                # If there is only one name, leave it unchanged
                formatted_name = name

            directors.append(formatted_name)

    return ", ".join(directors)


# ============================================================
# COUNTRY
# ============================================================

def get_countries(details):
    """
    Extract production countries.
    """

    countries = details.get("production_countries", [])

    return ", ".join(
        country.get("name", "")
        for country in countries
        if country.get("name")
    )


# ============================================================
# GENRES
# ============================================================

def get_genres(details):
    """
    Extract movie genres as a list.

    Example:
        ["Action", "Science Fiction"]
    """

    genres = details.get("genres", [])

    return [
        genre.get("name", "").strip()
        for genre in genres
        if genre.get("name")
    ]


# ============================================================
# READ TXT FILE
# ============================================================

def read_movies(filename):
    """
    Read movies from the TXT file.

    Expected format:

        Movie Title 2020
        Another Movie 1998

    The last 4 digits of each line are treated as the
    movie's release year. Everything before the year
    is treated as the movie title.

    Duplicate movies with the same title and year
    are automatically removed.
    """

    movies = []

    # Keep track of movies we've already seen.
    # Each entry will be:
    # ("movie title", year)
    seen_movies = set()

    with open(
        filename,
        "r",
        encoding="utf-8"
    ) as file:

        for line_number, line in enumerate(
            file,
            start=1
        ):

            line = line.strip()

            # Ignore empty lines
            if not line:
                continue

            # Ignore comments
            if line.startswith("#"):
                continue

            # Make sure the line has at least 4 characters
            if len(line) < 4:

                print(
                    f"WARNING: Line {line_number} skipped "
                    f"(could not find a 4-digit year)"
                )

                continue

            # Last 4 characters are the year
            year_text = line[-4:]

            # Everything before the year is the title
            title = line[:-4].strip()

            # Check that the year is numeric
            if not year_text.isdigit():

                print(
                    f"WARNING: Line {line_number} skipped "
                    f"(last 4 characters are not a year: "
                    f"{year_text})"
                )

                continue

            year = int(year_text)

            # Basic year validation
            if year < 1800 or year > 2100:

                print(
                    f"WARNING: Line {line_number} skipped "
                    f"(invalid year: {year})"
                )

                continue

            # Make sure a title exists
            if not title:

                print(
                    f"WARNING: Line {line_number} skipped "
                    f"(movie title is missing)"
                )

                continue

            # ------------------------------------------------
            # CHECK FOR DUPLICATES
            # ------------------------------------------------

            # Normalize the title so that these are considered
            # the same movie:
            #
            # The Godfather 1972
            # the godfather 1972
            
            normalized_title = title.lower().strip()

            movie_key = (
                normalized_title,
                year
            )

            if movie_key in seen_movies:

                print(
                    f"Duplicate skipped: "
                    f"{title} ({year})"
                )

                continue

            # Add movie to our set of seen movies
            seen_movies.add(movie_key)

            # Add movie to the list
            movies.append({
                "input_title": title,
                "input_year": year
            })

    return movies


# ============================================================
# PROCESS MOVIES
# ============================================================

def process_movies(movies):

    rows = []

    # Keep track of every genre encountered.
    # A set prevents duplicates, while the list preserves
    # the order in which genres were first encountered.
    all_genres = []
    genre_set = set()

    total = len(movies)

    for index, movie in enumerate(movies, start=1):

        input_title = movie["input_title"]
        input_year = movie["input_year"]

        print(
            f"[{index}/{total}] Searching: "
            f"{input_title} ({input_year})"
        )

        try:
            # Search TMDB
            search_result = search_movie(
                input_title,
                input_year
            )

            if search_result is None:

                print("    -> NOT FOUND")

                rows.append({
                    "Director": "",
                    "Title": input_title,
                    "Year": input_year,
                    "Country": "",
                    "Genres": [],
                    "_genres": []
                })

                continue

            movie_id = search_result["id"]

            # Retrieve complete details
            details = get_movie_details(movie_id)

            # Use the actual TMDB title/year
            title = details.get(
                "title",
                input_title
            )

            release_date = details.get(
                "release_date",
                ""
            )

            year = (
                release_date[:4]
                if release_date
                else input_year
            )

            director = get_director(details)
            countries = get_countries(details)
            genres = get_genres(details)

            # Add any newly encountered genres
            for genre in genres:

                if genre not in genre_set:
                    genre_set.add(genre)
                    all_genres.append(genre)

            rows.append({
                "Director": director,
                "Title": title,
                "Year": year,
                "Country": countries,
                "Genres": genres,
                "_genres": genres
            })

            print(
                f"    -> {director} | "
                f"{title} | "
                f"{year} | "
                f"{countries} | "
                f"{', '.join(genres)}"
            )

        except requests.exceptions.HTTPError as e:

            print(f"    -> HTTP ERROR: {e}")

            rows.append({
                "Director": "",
                "Title": input_title,
                "Year": input_year,
                "Country": "",
                "Genres": [],
                "_genres": []
            })

        except requests.exceptions.RequestException as e:

            print(f"    -> REQUEST ERROR: {e}")

            rows.append({
                "Director": "",
                "Title": input_title,
                "Year": input_year,
                "Country": "",
                "Genres": [],
                "_genres": []
            })

        except Exception as e:

            print(f"    -> ERROR: {e}")

            rows.append({
                "Director": "",
                "Title": input_title,
                "Year": input_year,
                "Country": "",
                "Genres": [],
                "_genres": []
            })

        # Small delay between requests
        time.sleep(0.25)

    return rows, all_genres


# ============================================================
# WRITE EXCEL
# ============================================================

def save_to_excel(rows, all_genres, filename):
    # --------------------------------------------------------
    # Create workbook
    # --------------------------------------------------------
    workbook = xlsxwriter.Workbook(filename)

    worksheet = workbook.add_worksheet("Movies")

    # --------------------------------------------------------
    # FORMATS
    # --------------------------------------------------------

    header_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    cell_format = workbook.add_format({
        "border": 1,
        "valign": "vcenter"
    })

    year_format = workbook.add_format({
        "border": 1,
        "align": "center"
    })

    # --------------------------------------------------------
    # FIXED COLUMNS
    # --------------------------------------------------------

    fixed_columns = [
        "Director",
        "Title",
        "Year",
        "Country",
        "Genres"
    ]

    # Write fixed column headers
    for col, column_name in enumerate(fixed_columns):

        worksheet.write(
            0,
            col,
            column_name,
            header_format
        )

    # --------------------------------------------------------
    # GENRE COLUMNS
    # --------------------------------------------------------

    genre_start_col = len(fixed_columns)

    for genre_index, genre in enumerate(all_genres):

        col = genre_start_col + genre_index

        # Header = genre name
        worksheet.write(
            0,
            col,
            genre,
            header_format
        )

    # --------------------------------------------------------
    # WRITE MOVIE ROWS
    # --------------------------------------------------------

    for row_index, movie in enumerate(rows, start=1):

        # --------------------------------------------
        # Director
        # --------------------------------------------
        
        worksheet.write(
            row_index,
            0,
            movie["Director"],
            cell_format
        )

        # --------------------------------------------
        # Title
        # --------------------------------------------

        worksheet.write(
            row_index,
            1,
            movie["Title"],
            cell_format
        )

        # --------------------------------------------
        # Year
        # --------------------------------------------

        worksheet.write(
            row_index,
            2,
            movie["Year"],
            year_format
        )

        # --------------------------------------------
        # Country
        # --------------------------------------------

        worksheet.write(
            row_index,
            3,
            movie["Country"],
            cell_format
        )

        # --------------------------------------------
        # Genres
        # --------------------------------------------

        worksheet.write(
            row_index,
            4,
            ", ".join(movie["_genres"]),
            cell_format
        )

        # ----------------------------------------------------
        # GENRE CHECKBOXES
        # ----------------------------------------------------

        movie_genres = set(movie["_genres"])

        for genre_index, genre in enumerate(all_genres):

            col = genre_start_col + genre_index

            # TRUE if this movie has the genre
            has_genre = genre in movie_genres

            # XlsxWriter checkbox
            worksheet.insert_checkbox(
                row_index,
                col,
                has_genre
            )

    # --------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------

    worksheet.set_column(
        "A:A",
        35
    )

    worksheet.set_column(
        "B:B",
        40
    )

    worksheet.set_column(
        "C:C",
        10
    )

    worksheet.set_column(
        "D:D",
        30
    )

    worksheet.set_column(
        "E:E",
        40
    )

    # Genre columns
    if all_genres:

        worksheet.set_column(
            genre_start_col,
            genre_start_col + len(all_genres) - 1,
            15
        )

    # --------------------------------------------------------
    # SHEET SETTINGS
    # --------------------------------------------------------

    # Freeze header row
    worksheet.freeze_panes(1, 0)

    # Autofilter
    last_col = genre_start_col + len(all_genres) - 1

    if last_col >= 0:

        worksheet.autofilter(
            0,
            0,
            len(rows),
            last_col
        )

    # Make the header row taller
    worksheet.set_row(0, 25)

    # Close workbook
    workbook.close()

# ============================================================
# MAIN
# ============================================================

def main():

    # Check API token
    if (
        not TMDB_TOKEN
        or TMDB_TOKEN == "YOUR_TMDB_BEARER_TOKEN"
    ):
        print(
            "ERROR: Please enter your TMDB API Read Access Token."
        )
        print(
            "Either edit TMDB_TOKEN in this script or set "
            "the TMDB_TOKEN environment variable."
        )
        sys.exit(1)

    # Check input file
    if not os.path.exists(INPUT_FILE):

        print(
            f"ERROR: Could not find '{INPUT_FILE}'."
        )

        print(
            "Create a TXT file with one movie per line, "
            "using: Movie Title | Year"
        )

        sys.exit(1)

    # Read movies
    movies = read_movies(INPUT_FILE)

    if not movies:
        print("ERROR: No valid movies found in the input file.")
        sys.exit(1)

    print(
        f"\nFound {len(movies)} movies.\n"
    )

    # Process
    rows, all_genres = process_movies(movies)

    # Save
    save_to_excel(
        rows,
        all_genres,
        OUTPUT_FILE
    )

    print(
        f"\nDone!"
    )

    print(
        f"Excel file created: {OUTPUT_FILE}"
    )


if __name__ == "__main__":
    main()