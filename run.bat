set classpath=lib/*;
javac -d build src/testcases/*.java src/utils/*.java
java -cp build;lib/* org.testng.TestNG testng.xml
