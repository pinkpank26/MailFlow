import tempfile
import unittest
from pathlib import Path
from mailer import EmailDataProcessor, SmartEmailSender


class MailerTests(unittest.TestCase):
    def test_extract_and_deduplicate(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'input.csv'
            path.write_text('email\nFIRST@example.com; second@example.org\nfirst@example.com\ninvalid\n')
            data = EmailDataProcessor(str(path))
            self.assertEqual(data.get_emails_list(), ['first@example.com', 'second@example.org'])

    def test_empty_input(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'input.csv'
            path.write_text('email\ninvalid\n')
            self.assertEqual(EmailDataProcessor(str(path)).get_emails_list(), [])

    def test_state_survives_restart_and_limits_apply(self):
        with tempfile.TemporaryDirectory() as directory:
            path = str(Path(directory) / 'state.json')
            accounts = [{'email': 'sender@example.com', 'daily_limit': 1}]
            sender = SmartEmailSender(accounts, path)
            sender.update_sent_count('sender@example.com')
            restored = SmartEmailSender(accounts, path)
            self.assertEqual(restored.calculate_remaining_capacity(), 0)
            self.assertIsNone(restored.get_available_email_config())


if __name__ == '__main__':
    unittest.main()
