from testcase import TestCase

class WasRun(TestCase):
    def __init__(self, name):
        self.wasRun = None
        super().__init__(name)

    def testMethod(self):
        self.wasRun = 1