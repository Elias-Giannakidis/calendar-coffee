class Person:
    def __init__(self, name, constrains, lastWeek):
        self.name = name
        self.constrains = constrains
        self.calendar = []
        self.initCalendar = []
        self.lastWeek = lastWeek
        self.MAX_WORK_DAYS = 6
        self.SCORE = {
            'weekend': 100,
            'program': 1000,
            'evening2morning': -10,
            'repo-motivo': 500,
            'repo': -200
        }
        self.repoWeek = 2

    def printMe(self):
        print('name = ', self.name)
        print('constrains = ', self.constrains)
        print('last week = ', self.lastWeek)
        print('calendar = ', self.calendar)
        print('-------------------------------------')

    def getRepoScore(self, calendar):
        score = 0
        for day in calendar:
            if day == 'repo':
                score = score + self.SCORE['repo']
        return score

    def getRepoMotivoScore(self, calendar):
        score = 0
        count = 0
        init = [0, 0, 0, 0, 0, 0, 0]
        week = init.copy()
        for day in calendar:
            if day == 'repo':
                week[count] = 1
            if count == 6:
                if self.motivoInclunted(week):
                    score = score + self.SCORE['repo-motivo']
                week = init.copy()
                count = -1
            count = count + 1
        return score


    def genRepoMotivo(self):
        buffer = []
        days = self.repoWeek
        init = [0, 0, 0, 0, 0, 0, 0]
        for i in range(8 - days):
            newMotivo = init.copy()
            for j in range(i, i + days):
                newMotivo[j] = 1
            buffer.append(newMotivo)
        return buffer

    def motivoInclunted(self, calendar):
        def check(motivo, calendar):
            j = 0
            for day in motivo:
                if day == 1 and calendar[j] == 0:
                    return False
                j = j + 1
            return True
        buffer = self.genRepoMotivo()
        for motivo in buffer:
            if check(motivo, calendar):
                return True
        return False

    def getCalendarValid(self, calendar):
        valid = True
        count = 0
        seri = 0
        for day in calendar:
            if (day == 'repo'):
                seri = 0
            else:
                seri = seri + 1
            weekday = self.calendar[count]['weekday']
            if weekday in self.constrains:
                constrain = self.constrains[weekday]
            else:
                if 'Everyday' in self.constrains:
                    constrain = self.constrains['Everyday']
                else:
                    constrain = 'all'
            constrainValid = constrain == 'all' or constrain == day or self.calendar[count]['value'] == day or day == 'repo' or day == None
            valid = valid and constrainValid and seri <= self.MAX_WORK_DAYS
            if not valid:
                return False
            count = count + 1
        return valid

    def getScore(self, calendar):
        score = 0
        score = score + self.getScoreWeekend(calendar)
        score = score + self.getScoreProgramm(calendar)
        score = score + self.getScoreEveningToMorning(calendar)
        score = score + self.getRepoMotivoScore(calendar)
        return score

    # Get a reward if the schedule goes with the employs program
    def getScoreProgramm(self, calendar):
        score = 0
        count = 0
        for day in calendar:
            if(self.calendar[count]['value'] == day):
                score = score + self.SCORE['program']
            count = count + 1
        return score

    # Get a negative reward if the employ works from evening to morning
    def getScoreEveningToMorning(self, calendar):
        lastDay = 'repo'
        score = 0
        for day in calendar:
            if(lastDay == 'evening' and day == 'morning'):
                score = score - self.SCORE['evening2morning']
            lastDay = day
        return score

    # get a reward if the employ has repo at the weekend
    def getScoreWeekend(self, calendar):
        seriWeedays = False
        count = 0
        score = 0
        for day in calendar:
            weekday = self.calendar[count]['weekday']
            if(day == 'repo' and (weekday == 'Saturday' or weekday == 'Sunday')):
                if(seriWeedays):
                    score = self.SCORE['weekend']
                    return score
                else:
                    seriWeedays = True
            else:
                seriWeedays = False
            count = count + 1
        return score



# ====== #
# =
# ====== #

# From staff_people get the people. It is un-efficient but who cares?
def genPeople(staff_people):
    const_keys = ['Everyday', 'Monday', 'Tuesday', 'wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    people = []
    for staff_person in staff_people:
        if "Name" in staff_person:
            newConstrains = {}
            for key in const_keys:
                if key in staff_person:
                    newConstrains[key] = staff_person[key]
            newPerson = Person(name=staff_person['Name'], constrains=newConstrains, lastWeek=[])
            people.append(newPerson)
    return people
