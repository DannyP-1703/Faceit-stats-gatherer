import requests
import bs4
import xl_funcs


def parse_users_data(username: str):

    URL = f'https://faceitelo.net/player/{username}'

    #       Downloading HTML
    r = requests.get(URL)
    file = r.text
    open("stats.html", 'w').write(file)

    #       Finding a required tag in the soup
    matches_html = bs4.BeautifulSoup(file, "html.parser").find('table', class_='table table-hover').find('tbody')

    #       Making the matrix with stats
    matches_list = []
    for m in matches_html.children:
        match_stats = []
        if isinstance(m, bs4.Tag):
            for s in m.stripped_strings:
                match_stats.append(repr(s).replace('\'', ''))
            matches_list.append(match_stats)

    #       Sorting the matrix
    matches_list = sorted(matches_list, key=lambda match: xl_funcs.strin_to_date(match[8]))

    return matches_list


def refresh_stats(workbook, username: str, player_id):
    sheet_name = f'{username}\'s stats'
    ml = parse_users_data(username)
    try:
        ws = workbook[sheet_name]
    except KeyError:
        ws = workbook.create_sheet(sheet_name)
        print(f"{player_id}. New worksheet created")
        xl_funcs.add_template(ws)
        print(f"{player_id}. Template is added")
    xl_funcs.fill_in_stats(ws, ml)
    print(f"{player_id}. Stats are filled in")
    xl_funcs.apply_styles(ws)
    print(f"{player_id}. Styles are applied")
