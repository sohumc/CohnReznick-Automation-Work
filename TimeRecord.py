__author__ = 'schitalia'

class TimeRecord(object):
    def __init__(self,date,hours,applicant,pwnum,task,taskdesc):
        self.date = date
        self.hours = hours
        self.applicant = applicant
        self.pwnum = pwnum
        self.task = task
        self.taskdesc = taskdesc

    def __eq__(self, other):
        if self.date != other.date or self.hours != other.hours or self.applicant != other.applicant or self.pwnum != other.pwnum or self.task != other.task or self.taskdesc != other.taskdesc:
            return False
        else:
            return True

    def __ne__(self, other):
        return not self.__eq__(other)

