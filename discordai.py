import discord
import activityspread
import asyncio


class MyClient(discord.Client):
    def __init__(self):
        discord.Client.__init__(self)
        self.spreadsheet_accessor = activityspread.SpreadsheetHandler()

        self.registered = {}
        self.registered_reverse = {}
        self.boss_queue = {}

        self.cra = ['cq', 'cp', 'cvb', 'cvel']
        self.boss_queue_index = {'cq': 1, 'cp': 4, 'cvb': 7, 'cvel': 10, 'hmag': 1, 'hellux': 4, 'hhilla': 7}
        self.boss_queue_names = {'cq': 'Chaos Crimson Queen',
                                 'cp': 'Chaos Pierre',
                                 'cvb': 'Chaos Von Bon',
                                 'cvel': 'Chaos Vellum',
                                 'hmag': 'Hard Magnus',
                                 'hellux': 'Hell Gollux',
                                 'hhilla': 'Hard Hilla'}

        self.oceaniaGuildActivity = self.spreadsheet_accessor.open_spreadsheet('oceaniaGuildActivity')
        self.oceaniaCarryQueue_sheet1 = self.spreadsheet_accessor.open_spreadsheet('oceaniaCarryQueue')
        self.oceaniaCarryQueue_sheet2 = self.spreadsheet_accessor.open_spreadsheet('oceaniaCarryQueue', worksheet=1)

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

        id_list = self.spreadsheet_accessor.get_column_values_raw(self.oceaniaGuildActivity, 4)

        for x, discord_id in enumerate(id_list):
            if discord_id is not '':
                self.registered[self.spreadsheet_accessor.get_value(self.oceaniaGuildActivity, x+1, 1)] = discord_id

        self.registered_reverse = {value: key for key, value in self.registered.items()}


    async def establish_connection(self):
        await self.wait_until_ready()
        while not self.is_closed:
            print('restarting loop...')
            await asyncio.sleep(300)
            print('reestablishing...')
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
            if discord_id in self.registered_reverse.keys():
                yield from self.send_message(message.channel, message.author.mention + ' You are already registered as {}.'.format(self.registered_reverse[discord_id]))
            else:
                yield from self.send_message(message.channel, message.author.mention + ' You are about to be registered as ' + name + '. Type *confirm to continue or *cancel to cancel registration.')
                yield from self.register(message)
        elif message.content.startswith('*carryme'):
            yield from self.send_message(message.channel, message.author.mention + ' Joining boss queue...')
            yield from self.carry_me(message)
        elif message.content.startswith('*weekly'):
            pass
        elif message.content.startswith('*update'):
            pass
            if message.author.top_role not in message.server.roles[:2]:
                yield from self.send_message(message.channel, message.author.mention + ' You do not have the required permissions to use this command.')
            else:
                yield from self.update(message)
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
        yield from self.send_message(message.channel, message.author.mention +
                                     ' If you would  like to add the names to the spreadsheet, use *confirm. '
                                     'If you would like to review the names, use *review, or *cancel to stop '
                                     'any further actions')
        reply = yield from self.wait_for_message(author=message.author)

        if reply.content.startswith('*confirm'):
            yield from self.update(missing_names, force=True)
        elif reply.content.startswith('*review'):
            yield from self.send_message(message.channel, message.author.mention +
                                         ' These were the names that were not recognized'
                                         ' to be on the spreadsheet:\n' + missing_names)
            yield from self.name_missing(message, missing_names)
        elif reply.content.startswith('*cancel'):
            yield from self.send_message(message.channel, message.author.mention + ' New name addition cancelled..')

    def _send(self, message, message_channel, mention=False, author=None):
        if mention:
            return self.send_message(message_channel, author.mention + message)
        else:
            return self.send_message(message_channel, message)

    @asyncio.coroutine
    def update(self, message, force=False):
        sent_message_generator = asyncio.gather(asyncio.ensure_future(self._send(' Updating spreadsheet...', message.channel, True, message.author)))
        sent_message = yield from sent_message_generator

        names = message.content[len('*update'):].replace(" ", "")
        name_list = sheet_object.spreadsheet_get_names('activity_sheet', 1, 1)

        not_in_list = sheet_object.spreadsheet_update_activity('activity_sheet', names, name_list, force)[:-2]

        if len(not_in_list) > 0:
            yield from self.edit_message(sent_message[0], message.author.mention +
                                         ' Spreadsheet has been updated. There were {} new name(s)'
                                         ' that were not previously on the spreadsheet.'
                                         .format(len(not_in_list.replace(' ', '').split(','))))

            yield from self.name_missing(message, not_in_list)
        else:
            yield from self.edit_message(sent_message[0], message.author.mention + ' Spreadsheet has been updated.')

        return False

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
        name_list = self.spreadsheet_accessor.get_column_values(self.oceaniaGuildActivity, 1)
        discord_id = str(message.author.id)
        self.check_if_registered(discord_id)
        if name in name_list:
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, self.spreadsheet_accessor.get_row(self.oceaniaGuildActivity, 1, name) + 1, 4, discord_id)
            self.add_to_memory(name, discord_id)
            yield from self.send_message(message.channel, message.author.mention + ' You have been successfully registered.')
        else:
            row = len(list(filter(None, name_list))) + 1
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 1, name)
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 2, name.upper())
            self.spreadsheet_accessor.write_to_spreadsheet(self.oceaniaGuildActivity, row, 4, discord_id)
            self.add_to_memory(name, discord_id)
            yield from self.send_message(message.channel, message.author.mention + ' You have been successfully registered.')

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

    @asyncio.coroutine
    def weekly(self, message):
        pass