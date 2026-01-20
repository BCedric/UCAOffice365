<?php

namespace BCedric\UCAOffice365;

use BCedric\UCAOffice365\Service\UCAOffice365;
use BCedric\UCAOffice365\Service\GraphAPI;
use Symfony\Component\Config\Definition\Configurator\DefinitionConfigurator;
use Symfony\Component\DependencyInjection\ContainerBuilder;
use Symfony\Component\DependencyInjection\Loader\Configurator\ContainerConfigurator;
use Symfony\Component\HttpKernel\Bundle\AbstractBundle;



class BCedricUCAOffice365Bundle extends AbstractBundle
{

    public function configure(DefinitionConfigurator $definition): void
    {
        
    }

    public function loadExtension(array $config, ContainerConfigurator $container, ContainerBuilder $builder): void
    {
        $container->import('../config/services.yaml');
    }
}
