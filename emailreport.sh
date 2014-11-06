#!/bin/bash
HOSTNAME=`hostname`
count=0
for i in `ls -at /var/www/html/rackxls/Blades*.txt | head -2| awk '{print $1}' 2> /dev/null`
do
if [ $count -eq 0 ]; then
 newbladelist=`echo $i`
 echo "New Blade List: $newbladelist"
fi
if [ $count -eq 1 ]; then
 oldbladelist=`echo $i`
 echo "Old Blade List: $oldbladelist"
fi
count=$(( ${count}+1 ));
done;
diff $oldbladelist $newbladelist > /var/www/html/rackxls/Blade_changes
if [[ ! -s /var/www/html/rackxls/Blade_changes ]] ; then
    exit 1
fi
NEWXLS=`ls -at /var/www/html/rackxls/Blades*.xls | head -1 | cut -d \/ -f 6-`
echo " " >> /var/www/html/rackxls/Blade_changes
echo "---------------------------" >> /var/www/html/rackxls/Blade_changes
echo " " >> /var/www/html/rackxls/Blade_changes
echo "New Chassis XLS can be found here: http://$HOSTNAME/rackxls/$NEWXLS" >> /var/www/html/rackxls/Blade_changes
mail -s "[Blade Moves / Changes] - Reith" blah@blah.com < /var/www/html/rackxls/Blade_changes  -- -f BladeReport
