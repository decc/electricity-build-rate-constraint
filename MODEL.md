# Theory


## Demand

Divide demand into 1 GW units. Each unit can be ON or OFF.

Each unit has a load factor. The load factor indicates the number of half hours in a year that a unit is ON in a year.

Some units will be on continuously. They have a load factor of 100%. Some units will never be on. They have a load factor of 0%.

## Supply

Divide supply into three groups:

1. Intermittent low carbon (i.e., wind)
2. Dispatchable low carbon (i.e., nuclear or CCS)
3. Dispatchable high carbon (i.e., CCGT gas)

### Dispatchable low carbon and Dispatchable high carbon

We divide dispatchable generation into a series of 1 GW units. Each unit may be ON or OFF.  Each unit has an availability factor. The availability factor indicates the number of half hours in a year that a unit could be ON, if there was demand. For dispatchable low carbon and high carbon, each unit would have an availability factor close to 100% -- if they are needed, they are available.

Each unit has a load factor. This is the availability factor multiplied by the proportion of time that there was demand for that unit. i.e., it is the proportion of the year that the unit could be ON, and in fact is ON, because demand warrants it.


### Intermittent low carbon

For intermittent low carbon we imagine that wind comes in 'puffs' and that wind turbines are arranged in a line, one behind each other. A light puff would cause the front most turbine to turn, but not the ones behind it. A stronger puff would cause the front turbine AND the one behind it to spin. Stronger still and the first three turbines might spin. The strongest possible puff would cause all the turbines to spin at the same time. The frequency of different strengths would be assumed to follow a (simplification of the) Weibull distribution. This would lead the front most turbine to spin almost all the time, the back most turbine to spin very rarely. In this way, we might assume that wind can be represented as a series of 1 GW units, some of which are ON close to 100% of the time and some of which are ON almost 0% of the time. Others are ON fractions inbetween. This means we have a situation like dispatchable generation, but where the availability factor varies by unit. 

Each unit will have a load factor. As with dispatchable generation, this is the availability factor of the unit multiplied by the proportion of time that there was demand for that unit. i.e., the proprtion of the year that there is both enough puff to turn that unit AND there is enough demand to want the electricity it produces.

## Matching Supply and Demand

The assumption is that, for a given unit of demand, it is preferable to meet it with supply from intermittent low carbon first, if that low carbon is available, then from dispatchable low carbon second, if dispatchable low carbon is available, then dispatchable high carbon if nothing else is available.

In practice this means, for each 1 GW unit of demand, matching one or 1 GW units of supply to it.

The principal we use is to match the units of demand that are ON the most frequently with the units of supply that are available the most frequently.

Step 1 is to start with the unit of demand that is ON the most often, and match it with the unit of intermittent wind that is on the most often. We assume that their probability of being ON is not correlated. This means that if a unit of demand is ON 100% of the time, and the most frequent 

[FLESH THIS OUT]

Load factor wind

Residual load factor

For each dispatchable low carbon, in desencing order of availability factor, pick highest remaining residual load factor demand unit, multiply togehter to get that dispatchable unit's load factor.

Then, for each dispatchable high carbon, in descending order of availability factor, pick the highest remaining residual load factor demand unit, and multiply together that dispatchable unit's load factor.

Done.
