import argparse
from src.link_with_file_path import FileParser

parser = argparse.ArgumentParser(description='Scrape data from CapCut static HTML file and dump to Excel')

parser.add_argument('-p', '--html_path', type=str, help='Path to the HTML file', required=True)
parser.add_argument('-lp', '--local_resource_path', type=str, help='Path to the file sources', required=True)
parser.add_argument('-o', '--output_file_path', type=str, help='Path to the output file; default: output.xlsx', required=False)
parser.add_argument('-wi', '--workspace_id', type=str, help='Workspace ID; default: 7293489460916830210', required=False)

args = parser.parse_args()
if __name__ == '__main__':
    # soup = parse_html()
    # data = extract_data(soup)
    # dump_excel(data)
    if args.html_path:
        print(f'HTML file path: {args.html_path}')

    if args.local_resource_path:
        print(f'Local resource path: {args.local_resource_path}')

    if args.workspace_id:
        print(f'Workspace ID: {args.workspace_id}')
    else:
        print('Using default Workspace ID: 7293489460916830210')

    fp = FileParser(args.html_path, args.local_resource_path, args.output_file_path, args.workspace_id)
    fp.run()
    print('Done')
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
