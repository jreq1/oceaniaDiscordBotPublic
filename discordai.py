import discord
import activityspread
import asyncio
from datetime import datetime
import threading
from queue import Queue


class MyClient(discord.Client):
    def __init__(self):
        discord.Client.__init__(self)
        self.spreadsheet_accessor = activityspread.SpreadsheetHandler()

        self.registered = {}
        self.registered_reverse = {}
        self.boss_queue = {}

        self.cra = ['cq', 'cp', 'cvb', 'cvel']
        self.not_in_list = []

        self.boss_queue_index = {'cq': 1, 'cp': 4, 'cvb': 7, 'cvel': 10, 'hmag': 1, 'hellux': 4, 'hhilla': 7}
        self.boss_queue_names = {'cq': 'Chaos Crimson Queen',
                                 'cp': 'Chaos Pierre',
                                 'cvb': 'Chaos Von Bon',
                                 'cvel': 'Chaos Vellum',
                                 'hmag': 'Hard Magnus',
                                 'hellux': 'Hell Gollux',
                                 'hhilla': 'Hard Hilla'}

        self.list_of_names = None
        self.oceaniaGuildActivity = self.spreadsheet_accessor.open_spreadsheet('oceaniaGuildActivity')
        self.oceaniaCarryQueue_sheet1 = self.spreadsheet_accessor.open_spreadsheet('oceaniaCarryQueue')
        self.oceaniaCarryQueue_sheet2 = self.spreadsheet_accessor.open_spreadsheet('oceaniaCarryQueue', worksheet=1)

        self.q = Queue()
        for x in range(4):
            t = threading.Thread(target=self.worker)
            t.daemon = True
            t.start()

        self.loop.create_task(self.startup())
        self.refresher = self.loop.create_task(self.establish_connection())

    async def startup(self):
        self.registered = {}
        self.registered_reverse = {}
        self.boss_queue['cq'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet2, 1)
        self.boss_queue['cp'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet2, 4)
        self.boss_queue['cvb'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet2, 7)
        self.boss_queue['cvel'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet2, 10)
        self.boss_queue['hmag'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet1, 1)
        self.boss_queue['hellux'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet1, 4)
        self.boss_queue['hhilla'] = self.spreadsheet_accessor.get_column_values(self.oceaniaCarryQueue_sheet1, 7)
        self.list_of_names = self.spreadsheet_accessor.get_column_values(self.oceaniaGuildActivity, 7)

        id_list = self.spreadsheet_accessor.get_column_values_raw(self.oceaniaGuildActivity, 4)

        for x, discord_id in enumerate(id_list):
            if discord_id is not '':
                self.registered[self.spreadsheet_accessor.get_value(self.oceaniaGuildActivity, x+1, 1)] = discord_id

        self.registered_reverse = {value: key for key, value in self.registered.items()}

    def worker(self):
        while True:
            items = self.q.get()
            self.write_update(items[0], items[1], items[2], items[3], items[4])
            self.q.task_done()

    @asyncio.coroutine
    def on_ready(self):
        print('Ready')

    @asyncio.coroutine
    def establish_connection(self):
        yield from self.wait_until_ready()
        while not self.is_closed:
            yield from asyncio.sleep(600)
            self.oceaniaGuildActivity = self.spreadsheet_accessor.open_spreadsheet('oceaniaGuildActivity')
            self.oceaniaCarryQueue_sheet1 = self.spreadsheet_accessor.open_spreadsheet('oceaniaCarryQueue')
            self.oceaniaCarryQueue_sheet2 = self.spreadsheet_accessor.open_spreadsheet('oceaniaCarryQueue', worksheet=1)

    @asyncio.coroutine
    def on_message(self, message):
        if message.content.startswith('*test'):
            yield from self.send_message(message.channel, message.author.mention + ' Hello!')
        elif message.content.startswith('*register'):
            name = message.content[len('*register'):].replace(' ', '')
            discord_id = str(message.author.id)
            if len(name) > 0:
                if discord_id in self.registered_reverse.keys():
                    yield from self.send_message(message.channel, message.author.mention + ' You are already registered as {}.'.format(self.registered_reverse[discord_id]))
                else:
                    yield from self.send_message(message.channel, message.author.mention + ' You are about to be registered as ' + name + '. Type *confirm to continue or *cancel to cancel registration.')
                    self.loop.create_task(self.register(message))
            else:
                yield from self.send_message(message.channel, message.author.mention + ' You need to enter your maple IGN to register.')
        elif message.content.startswith('*carryme'):
            if self.check_if_registered(str(message.author.id)):
                yield from self.carry_me(message)
            else:
                yield from self.send_message(message.channel, message.author.mention + ' You need to register first! Use *register (MapleIGN) to link your discord ID with your maple IGN.'
                                                                                       '\n Warning: You can only register your discord ID to one IGN.')
        elif message.content.startswith('*update'):
            if self.has_permission(message):
                self.loop.create_task(self.update(message))
        elif message.content.startswith('*refresh'):
            if self.has_permission(message):
                self.loop.create_task(self.startup())

    def has_permission(self, message):
        if message.author.top_role not in message.server.roles[:2]:
            yield from self.send_message(message.channel, message.author.mention + ' You do not have the required permissions to use this command.')
            return False
        else:
            return True

    @asyncio.coroutine
    def name_missing(self, message, missing_names):
        reply = yield from self.wait_for_message(author=message.author)
        if reply.content.startswith('*confirm'):
            self.loop.create_task(self.update(message, missing_names=missing_names, force=True))
        elif reply.content.startswith('*review'):
            yield from self.send_message(message.channel, message.author.mention +
                                         ' These were the names that were not recognized'
                                         ' to be on the spreadsheet:\n{}'.format(', '.join(missing_names)))
            yield from self.name_missing(message, missing_names)
        elif reply.content.startswith('*cancel'):
            yield from self.send_message(message.channel, message.author.mention + ' New name addition cancelled.')

    def _send(self, message, message_channel, mention=False, author=None):
        if mention:
            return self.send_message(message_channel, author.mention + message)
        else:
            return self.send_message(message_channel, message)

    def get_row(self, index):
        try:
            return self.list_of_names.index(index)
        except ValueError:
            return None

    @staticmethod
    def write_update(spreadsheet_accessor, spreadsheet, row, column, name):
        spreadsheet_accessor.write_to_spreadsheet(spreadsheet, row, column, name)

    @asyncio.coroutine
    def update(self, message, missing_names=None, force=False):
        sent_message_generator = asyncio.gather(asyncio.ensure_future(self._send(' Updating spreadsheet...', message.channel, True, message.author)))
        sent_message = yield from sent_message_generator

        self.not_in_list = []
        if message is None:
            names = missing_names
        else:
            names = message.content[len('*update'):].replace(" ", "").split(',')

        today = '{}/{}/{}'.format(datetime.now().day, datetime.now().month, datetime.now().year)
        last_value = len(self.spreadsheet_accessor.get_column_values(self.oceaniaGuildActivity, 1)) + 1
        for name in names:
            value = self.get_row(name.upper())
            if value is None:
                if force:
                    self.q.put([self.spreadsheet_accessor, self.oceaniaGuildActivity, last_value, 6, name])
                    self.q.put([self.spreadsheet_accessor, self.oceaniaGuildActivity, last_value, 7, name.upper()])
                    self.q.put([self.spreadsheet_accessor, self.oceaniaGuildActivity, last_value, 8, today])
                    last_value += 1
                    # self.loop.create_task(self.write_update(last_value, 6, name))
                    # self.loop.create_task(self.write_update(last_value, 7, name.upper()))
                    # self.loop.create_task(self.write_update(last_value, 8, today))
                else:
                    self.not_in_list.append(name)
            else:
                self.q.put([self.spreadsheet_accessor, self.oceaniaGuildActivity, value + 1, 8, today])

        new_message = ' Spreadsheet has been updated.'
        if len(self.not_in_list) > 0:
            new_message += ' There were {} new name(s) that were not previously on the spreadsheet. ' \
                       '\n If you would  like to add the names to the spreadsheet, use *confirm. ' \
                       'If you would like to review the names, use *review, or *cancel to stop any further actions'.format(len(self.not_in_list))
            yield from self.edit_message(sent_message[0], message.author.mention + new_message)
            self.loop.create_task(self.name_missing(message, self.not_in_list))
        else:
            yield from self.edit_message(sent_message[0], message.author.mention + new_message)

    # don't need to touch
    @staticmethod
    def check(message):
        if message.content.startswith('*confirm') or message.content.startswith('*cancel'):
            return True
        return False

    # done reformatting
    @asyncio.coroutine
    def register(self, message):
        reply = yield from self.wait_for_message(author=message.author)
        if self.check(reply):
            if reply.content.startswith('*confirm'):
                yield from self.send_message(message.channel, message.author.mention + ' You have been successfully registered.')
                yield from self.register_confirmed(message)
            else:
                yield from self.send_message(message.channel, message.author.mention + ' Cancelling registration.')
        else:
            yield from self.register(message)

    def add_to_memory(self, name, discord_id):
        self.registered[name] = discord_id
        self.registered_reverse[discord_id] = name

    # done reformatting
    @asyncio.coroutine
    def register_confirmed(self, message):
        name = message.content[len('*register'):].replace(' ', '')
        name_list = [x.lower() for x in self.spreadsheet_accessor.get_column_values(self.oceaniaGuildActivity, 1)]
        discord_id = str(message.author.id)
        self.check_if_registered(discord_id)
        if name.lower() in name_list:
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, self.spreadsheet_accessor.get_row(self.oceaniaGuildActivity, 2, name.upper()) + 1, 4, discord_id)
            self.add_to_memory(name, discord_id)
        else:
            row = len(list(filter(None, name_list))) + 1
            today = '{}/{}/{}'.format(datetime.now().day, datetime.now().month, datetime.now().year)
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 1, name)
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 2, name.upper())
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 4, discord_id)
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 8, today)
            self.add_to_memory(name, discord_id)

    # done reformatting
    def check_if_registered(self, discord_id):
        return discord_id in self.spreadsheet_accessor.get_column_values(self.oceaniaGuildActivity, 4)

    @asyncio.coroutine
    def write_boss_queue(self, message, boss):
        if boss in self.cra:
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaCarryQueue_sheet2, len(self.boss_queue[boss]), self.boss_queue_index[boss], self.registered_reverse[str(message.author.id)])
        else:
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaCarryQueue_sheet1, len(self.boss_queue[boss]), self.boss_queue_index[boss], self.registered_reverse[str(message.author.id)])

    @asyncio.coroutine
    def carry_me(self, message):
        bosses = message.content[len('*carryme'):].replace(' ', '').split(',')
        already_in_queue = ''
        not_in_queue = ''
        for boss in bosses:
            if self.registered_reverse[str(message.author.id)] not in self.boss_queue[boss]:
                not_in_queue += self.boss_queue_names[boss] + ', '
                self.boss_queue[boss].append(self.registered_reverse[str(message.author.id)])
                self.loop.create_task(self.write_boss_queue(message, boss))
            else:
                already_in_queue += self.boss_queue_names[boss] + ', '
        if len(already_in_queue) > 0:
            yield from self.send_message(message.channel, message.author.mention + ' You are already in the boss queue for ' + already_in_queue[:-2])
        if len(not_in_queue) > 0:
            yield from self.send_message(message.channel, message.author.mention + ' You have joined the boss queue for ' + not_in_queue[:-2])