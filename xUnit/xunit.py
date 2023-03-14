from wasrun import WasRun
from testcase import TestCase

test = WasRun("testMethod")
print(test.wasRun)
test.run()
print(test.wasRun)

