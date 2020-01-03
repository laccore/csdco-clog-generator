import argparse
import mailbox
import csv
import timeit
from gooey import Gooey, GooeyParser

def export_mbox(mbox_filename, output_filename):
  with open(output_filename, 'w') as out_file:
    writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)

    count = 0
    for message in mailbox.mbox(mbox_filename):
      count += 1
      data = [
        message['subject'],
        message['from'],
        message['to'],
        message['date']
      ]

      writer.writerow(data)

      if (count%100 == 0):
        print(f'\t{count} emails processed')

  return count

@Gooey(program_name='CSDCO CLOG Generator')
def main():
  parser = GooeyParser(description='Export data (Subject, From, To, Date) from a .mbox file to a CSV')
  parser.add_argument('mbox', metavar='.mbox file', widget='FileChooser', type=str, help='Name of mbox file')
  parser.add_argument('-o', '--output-filename', metavar='Output filename', type=str, help='Filename for export.')
  args = parser.parse_args()

  start_time = timeit.default_timer()

  mailbox_filename = args.mbox
  if 'output_filename' in args and args.output_filename:
    output_filename = args.output_filename
  else:
    output_filename = mailbox_filename.replace('.mbox', '.csv')

  # Export data
  print(f'Beginning export of {mailbox_filename} to {output_filename}...')

  message_count = export_mbox(mailbox_filename, output_filename)

  print(f'{message_count} emails were found and exported to {output_filename}.')
  print(f'Completed in {round((timeit.default_timer()-start_time), 2)} seconds.')


if __name__ == '__main__':
  main()
