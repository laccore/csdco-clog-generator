import mailbox
import csv
import timeit
import arrow
from gooey import Gooey, GooeyParser
from email.header import decode_header, make_header
import locale
import re
import os.path


def clean_header(header, verbose=False):
  try:
    return str(make_header(decode_header(re.sub(r'\s\s+', ' ', header)))) if header else ''
  except:
    return header


def process_mbox(mbox_filename, year=None, verbose=False):
  count = 0
  ignored = 0
  emails = []

  if verbose and year:
    print(f'Ignoring emails not from year {year}.')

  for message in mailbox.mbox(mbox_filename):
    count += 1
    
    # When downloading from Google Takeout, there are a few different datetime formats
    # So far, they've all matched one of the below options (with lots of regex for extra whitespace...)
    date_formats = [
      r'ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]Z',
      r'ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]ZZZ',
      r'ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]',
      r'ddd,[\s+]DD[\s+]MMM[\s+]YYYY[\s+]HH:mm:ss',
      r'ddd[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]Z',
      r'D[\s+]MMM[\s+]YYYY[\s+]HH:mm:ss[\s+]Z'
    ]

    a_date = None

    for d_format in date_formats:
      try:
        a_date = arrow.get(message['Date'], d_format)
        break
      except:
        continue
    
    if not a_date:
      print(f"ALERT: '{message['Date']}' does not match any expected format. Ignoring email with subject '{message['Subject']}'.")
      ignored += 1

    else:
      if year and (a_date.format('YYYY') != year):
        ignored += 1
        if verbose:
          print(f"WARNING: Invalid year found ({a_date.format('YYYY')}).")
      
      else:
        data = [
          clean_header(message['Subject']),
          clean_header(message['From']),
          clean_header(message['To']),
          a_date
        ]

        emails.append(data)

    if verbose and (count%1000 == 0):
      print(f'INFO: {count} emails processed.')

  # Sort based on arrow object
  emails = sorted(emails, key=lambda x: x[-1])

  # Convert list to desired string format
  date_format = 'M/D/YY'
  emails = [[*email[:-1], email[-1].format(date_format)] for email in emails]

  return [emails, count, ignored]


def export_emails(emails, output_filename):
  with open(output_filename, 'w', newline='', encoding=locale.getpreferredencoding()) as out_file:
    writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)
    writer.writerows(emails)


@Gooey(program_name='CSDCO CLOG Generator')
def main():
  parser = GooeyParser(description='Export data (Subject, From, To, Date) from a .mbox file to a CSV')
  parser.add_argument('mbox', metavar='.mbox file', widget='FileChooser', type=str, help='Name of mbox file')
  parser.add_argument('-o', '--output-filename', metavar='Output filename', type=str, help='Filename for export.')
  parser.add_argument('-y', '--year', metavar='Year', type=str, help='Ignore emails not from this year.')
  parser.add_argument('-v', '--verbose', metavar='Verbose', action='store_true', help='Print troubleshooting information.')
  args = parser.parse_args()

  start_time = timeit.default_timer()

  mailbox_filename = args.mbox
  if 'output_filename' in args and args.output_filename:
    if not os.path.exists(args.output_filename):
      if args.verbose:
        print(f"Invalid path name for export given. Changed to {mailbox_filename.replace('.mbox', '.csv')}")
      output_filename = mailbox_filename.replace('.mbox', '.csv')
    else:
      output_filename = args.output_filename
  else:
    output_filename = mailbox_filename.replace('.mbox', '.csv')

  # Process data
  print(f'Beginning processing of {mailbox_filename}...')
  emails, message_count, ignored_count = process_mbox(mailbox_filename, args.year, args.verbose)

  # Export data
  print(f'Beginning export of {len(emails)} emails to {output_filename}...')
  export_emails(emails, output_filename)

  print(f'{message_count} emails were found and {message_count - ignored_count} were exported to {output_filename}.')
  print(f'Completed in {round((timeit.default_timer()-start_time), 2)} seconds.')


if __name__ == '__main__':
  main()
